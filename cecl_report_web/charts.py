"""Read embedded workbook charts and re-render them as inline SVG.

openpyxl exposes each worksheet's charts (``ws._charts``) with their type,
title, and data/category references. We resolve those references to real
values and draw dependency-free, vector SVG — deterministic (good for the
visual-regression harness) and crisp in the PDF. Excel-default palette so
the charts read as "the same chart".
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter

# Excel 2016+ default series palette.
PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]

_REF_RE = re.compile(r"^(?:'([^']+)'|([^!]+))!(.+)$")


def _split_ref(ref: str) -> tuple[str | None, str]:
    """'Sheet'!$A$1:$B$2 -> ('Sheet', '$A$1:$B$2')."""
    m = _REF_RE.match(ref or "")
    if not m:
        return None, ref
    sheet = m.group(1) or m.group(2)
    return sheet, m.group(3)


def _resolve(wb, ref: str | None) -> list[Any]:
    """Resolve a cell/range reference to a flat list of values."""
    if not ref:
        return []
    sheet, a1 = _split_ref(ref)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    a1 = a1.replace("$", "")
    try:
        min_c, min_r, max_c, max_r = range_boundaries(a1)
    except Exception:  # noqa: BLE001
        return []
    out: list[Any] = []
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            out.append(ws.cell(r, c).value)
    return out


def _resolve_scalar(wb, ref: str | None) -> Any:
    vals = _resolve(wb, ref)
    return vals[0] if vals else None


def _num(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _chart_title(ch) -> str | None:
    try:
        if ch.title and ch.title.tx and ch.title.tx.rich:
            return "".join(
                (r.t or "") for p in ch.title.tx.rich.p for r in (p.r or [])
            ).strip() or None
    except Exception:  # noqa: BLE001
        pass
    return None


def read_chart_specs(report_path: str | Path, sheet: str) -> list[dict]:
    """Return normalized chart specs for one worksheet."""
    wb = load_workbook(report_path)
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    specs: list[dict] = []
    for ch in getattr(ws, "_charts", []) or []:
        ctype = type(ch).__name__
        series: list[dict] = []
        for s in getattr(ch, "series", []) or []:
            name = None
            try:
                if s.tx and s.tx.strRef:
                    name = _resolve_scalar(wb, s.tx.strRef.f)
                elif s.tx and s.tx.v:
                    name = s.tx.v
            except Exception:  # noqa: BLE001
                name = None
            vref = None
            try:
                vref = s.val.numRef.f if (s.val and s.val.numRef) else None
            except Exception:  # noqa: BLE001
                vref = None
            cref = None
            try:
                if s.cat and s.cat.numRef:
                    cref = s.cat.numRef.f
                elif s.cat and s.cat.strRef:
                    cref = s.cat.strRef.f
            except Exception:  # noqa: BLE001
                cref = None
            series.append({
                "name": name,
                "values": [_num(v) for v in _resolve(wb, vref)],
                "cats": [("" if v is None else str(v)) for v in _resolve(wb, cref)],
            })
        specs.append({
            "type": ctype,
            "bar_dir": getattr(ch, "type", None),      # 'col' | 'bar'
            "grouping": getattr(ch, "grouping", None),  # 'clustered' | 'stacked'
            "title": _chart_title(ch),
            "series": series,
        })
    return specs


# ── SVG rendering ────────────────────────────────────────────────────

import html as _html

_W, _H = 340, 240


def _esc(s) -> str:
    return _html.escape("" if s is None else str(s))


def _wrap(inner: str, title: str | None, w: int = _W, h: int = _H) -> str:
    t = (f'<text x="{w/2:.0f}" y="16" text-anchor="middle" '
         f'font-size="12" font-weight="700" fill="#000">{_esc(title)}</text>'
         if title else "")
    return (
        f'<svg class="chart" xmlns="http://www.w3.org/2000/svg" '
        f'viewBox="0 0 {w} {h}" width="{w}" height="{h}">'
        f'<rect x="0" y="0" width="{w}" height="{h}" fill="#fff"/>{t}{inner}</svg>'
    )


def _legend(items: list[tuple[str, str]], x: int, y: int) -> str:
    """items = [(label, color)]. Vertical legend."""
    out = []
    for i, (label, color) in enumerate(items):
        yy = y + i * 16
        out.append(
            f'<rect x="{x}" y="{yy}" width="10" height="10" fill="{color}"/>'
            f'<text x="{x + 14}" y="{yy + 9}" font-size="10" fill="#000">'
            f'{_esc(label)}</text>')
    return "".join(out)


def _svg_bar(series: list[dict], cats: list[str], title: str | None,
             *, horizontal: bool, stacked: bool) -> str:
    # Horizontal bars need a wider left gutter for category labels.
    top, right, bottom = 28, 12, 46
    left = 98 if horizontal else 46
    pw, ph = _W - left - right, _H - top - bottom
    nseries = max(1, len(series))
    multi = nseries > 1
    if multi:
        pw -= 70  # room for legend
    # Max magnitude across data (handle negatives for Net Change).
    all_vals = [v for s in series for v in s["values"]]
    if stacked:
        # per-category sum
        ncat = len(cats)
        sums = [sum(s["values"][i] for s in series if i < len(s["values"]))
                for i in range(ncat)]
        vmax = max([abs(x) for x in sums] + [0.0])
    else:
        vmax = max([abs(v) for v in all_vals] + [0.0])
    vmax = vmax or 1.0

    parts = [f'<rect x="{left}" y="{top}" width="{pw}" height="{ph}" '
             f'fill="none" stroke="#d9d9d9"/>']
    ncat = len(cats)
    group_w = pw / max(1, ncat)

    if horizontal:
        # categories stacked vertically; bars extend right.
        row_h = ph / max(1, ncat)
        bar_h = row_h * 0.55
        for i, cat in enumerate(cats):
            cy = top + i * row_h + (row_h - bar_h) / 2
            x0 = left
            if stacked:
                acc = 0.0
                for si, s in enumerate(series):
                    val = s["values"][i] if i < len(s["values"]) else 0.0
                    w = (val / vmax) * pw
                    parts.append(
                        f'<rect x="{x0 + acc:.1f}" y="{cy:.1f}" '
                        f'width="{max(0, w):.1f}" height="{bar_h:.1f}" '
                        f'fill="{PALETTE[si % len(PALETTE)]}"/>')
                    acc += w
            else:
                sh = bar_h / nseries
                for si, s in enumerate(series):
                    val = s["values"][i] if i < len(s["values"]) else 0.0
                    w = (val / vmax) * pw
                    parts.append(
                        f'<rect x="{x0:.1f}" y="{cy + si * sh:.1f}" '
                        f'width="{max(0, w):.1f}" height="{sh:.1f}" '
                        f'fill="{PALETTE[si % len(PALETTE)]}"/>')
            parts.append(
                f'<text x="{left - 4}" y="{top + i * row_h + row_h/2 + 3:.1f}" '
                f'text-anchor="end" font-size="8" fill="#000">{_esc(cat)}</text>')
    else:
        bw = group_w * 0.7 / nseries
        for i, cat in enumerate(cats):
            gx = left + i * group_w + (group_w - bw * nseries) / 2
            for si, s in enumerate(series):
                val = s["values"][i] if i < len(s["values"]) else 0.0
                bh = (abs(val) / vmax) * ph
                y = top + ph - bh
                parts.append(
                    f'<rect x="{gx + si * bw:.1f}" y="{y:.1f}" '
                    f'width="{bw:.1f}" height="{bh:.1f}" '
                    f'fill="{PALETTE[si % len(PALETTE)]}"/>')
            parts.append(
                f'<text x="{left + i * group_w + group_w/2:.1f}" '
                f'y="{top + ph + 12}" text-anchor="middle" font-size="8" '
                f'fill="#000">{_esc(cat)}</text>')

    if multi:
        parts.append(_legend(
            [(s["name"] or f"Series {i+1}", PALETTE[i % len(PALETTE)])
             for i, s in enumerate(series)],
            left + pw + 12, top + 4))
    return _wrap("".join(parts), title)


def _svg_pie(values: list[float], cats: list[str], title: str | None,
             *, doughnut: bool) -> str:
    import math
    cx, cy, rad = 100, 130, 78
    inner = rad * 0.55 if doughnut else 0
    total = sum(v for v in values if v > 0)
    parts: list[str] = []
    if total <= 0:
        parts.append(f'<circle cx="{cx}" cy="{cy}" r="{rad}" fill="none" '
                     f'stroke="#c9c9c9"/>'
                     f'<text x="{cx}" y="{cy}" text-anchor="middle" '
                     f'font-size="10" fill="#888">No data</text>')
    else:
        ang = -math.pi / 2
        for i, v in enumerate(values):
            if v <= 0:
                continue
            frac = v / total
            a2 = ang + frac * 2 * math.pi
            large = 1 if frac > 0.5 else 0
            x1, y1 = cx + rad * math.cos(ang), cy + rad * math.sin(ang)
            x2, y2 = cx + rad * math.cos(a2), cy + rad * math.sin(a2)
            color = PALETTE[i % len(PALETTE)]
            if frac >= 0.999:
                parts.append(f'<circle cx="{cx}" cy="{cy}" r="{rad}" '
                             f'fill="{color}"/>')
            else:
                parts.append(
                    f'<path d="M{cx},{cy} L{x1:.1f},{y1:.1f} '
                    f'A{rad},{rad} 0 {large} 1 {x2:.1f},{y2:.1f} Z" '
                    f'fill="{color}"/>')
            ang = a2
        if doughnut and inner:
            parts.append(f'<circle cx="{cx}" cy="{cy}" r="{inner:.0f}" '
                         f'fill="#fff"/>')
    parts.append(_legend(
        [(f"{cats[i] if i < len(cats) else ''} "
          f"({(values[i]/total*100 if total else 0):.0f}%)",
          PALETTE[i % len(PALETTE)])
         for i in range(len(values))],
        195, 100))
    return _wrap("".join(parts), title)


def render_chart_svg(spec: dict) -> str:
    """Dispatch a normalized chart spec to the right SVG renderer."""
    ctype = spec.get("type", "")
    series = spec.get("series", [])
    title = spec.get("title")
    if "Pie" in ctype or "Doughnut" in ctype:
        s0 = series[0] if series else {"values": [], "cats": []}
        return _svg_pie(s0.get("values", []), s0.get("cats", []), title,
                        doughnut="Doughnut" in ctype)
    # Bar/column chart.
    cats = series[0]["cats"] if series else []
    horizontal = spec.get("bar_dir") == "bar"
    stacked = spec.get("grouping") == "stacked"
    return _svg_bar(series, cats, title, horizontal=horizontal, stacked=stacked)


def render_charts_for_sheet(report_path: str | Path, sheet: str) -> list[str]:
    """Return a list of SVG strings for every chart on a worksheet."""
    return [render_chart_svg(s) for s in read_chart_specs(report_path, sheet)]

