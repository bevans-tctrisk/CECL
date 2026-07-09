"""Visual-regression harness for the browser report renderer.

Renders every tab of a set of fixture reports to deterministic PNGs and
compares them against stored reference images, so any unintended change
to the renderer (CSS, templates, formatters, chart SVG, ...) is caught
as a per-page pixel diff.

Usage is via ``scripts/report_visual_regression.py``:
    python scripts/report_visual_regression.py --update   # capture refs
    python scripts/report_visual_regression.py            # compare

Determinism: fixed viewport width per orientation, embedded fonts, print
media, ``document.fonts.ready`` wait. Same workbook + same code => same
pixels (within a small anti-aliasing tolerance).
"""

from __future__ import annotations

import io
import re
import tempfile
from pathlib import Path
from typing import Any

from . import assembly

# Viewport widths (px) — deterministic sizing; full-page capture drives height.
_W_PORTRAIT = 1100
_W_LANDSCAPE = 1500

# A pixel counts as "changed" when any channel differs by more than this.
_PIXEL_TOL = 20
# A page fails when more than this fraction of pixels changed. The renderer
# is deterministic (identical input -> identical pixels), so this can be
# tight; a small allowance absorbs any rare 1-2px anti-aliasing jitter.
_FAIL_FRACTION = 0.00005


def _safe(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")


def render_pages_png(report_path: str | Path, cu: str, snap: str,
                     *, supplemental: bool = False) -> dict[str, bytes]:
    """Render every worksheet of a report to PNG bytes, keyed by sheet."""
    from playwright.sync_api import sync_playwright

    pages = assembly.build_pages(report_path, cu, snap,
                                 supplemental=supplemental)
    out: dict[str, bytes] = {}
    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            for pg in pages:
                width = _W_LANDSCAPE if pg["landscape"] else _W_PORTRAIT
                page = browser.new_page(
                    viewport={"width": width, "height": 1200},
                    device_scale_factor=1)
                with tempfile.TemporaryDirectory() as td:
                    fp = Path(td) / "r.html"
                    fp.write_text(pg["full"], encoding="utf-8")
                    page.goto(fp.as_uri(), wait_until="load")
                    page.evaluate("document.fonts.ready")
                    page.emulate_media(media="print")
                    out[pg["sheet"]] = page.screenshot(full_page=True)
                page.close()
        finally:
            browser.close()
    return out


def _diff(ref_png: bytes, cur_png: bytes) -> tuple[float, bytes | None]:
    """Compare two PNGs. Returns (fraction_changed, diff_png_or_None).

    Dimension mismatch => fraction 1.0 (treated as a full change).
    """
    import numpy as np
    from PIL import Image

    ref = np.asarray(Image.open(io.BytesIO(ref_png)).convert("RGB"), dtype=np.int16)
    cur = np.asarray(Image.open(io.BytesIO(cur_png)).convert("RGB"), dtype=np.int16)
    if ref.shape != cur.shape:
        return 1.0, None
    delta = np.abs(ref - cur).max(axis=2)
    mask = delta > _PIXEL_TOL
    frac = float(mask.mean())
    if frac == 0.0:
        return 0.0, None
    # Diff image: dimmed current with changed pixels flagged red.
    dim = (cur * 0.35).astype(np.uint8)
    dim[mask] = [255, 0, 0]
    buf = io.BytesIO()
    Image.fromarray(dim, "RGB").save(buf, format="PNG")
    return frac, buf.getvalue()


def run(fixtures: list[dict[str, Any]], refs_dir: str | Path,
        *, update: bool = False,
        fail_fraction: float = _FAIL_FRACTION) -> dict[str, Any]:
    """Capture (``update``) or compare fixture renders against references.

    Returns a summary dict; ``ok`` is False if any page regressed.
    """
    refs_root = Path(refs_dir)
    refs_root.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    ok = True

    for fx in fixtures:
        label = fx["label"]
        report_path = fx["report_path"]
        if not Path(report_path).exists():
            results.append({"fixture": label, "status": "missing-report",
                            "report": str(report_path)})
            ok = False
            continue
        fx_dir = refs_root / _safe(label)
        fx_dir.mkdir(parents=True, exist_ok=True)
        pngs = render_pages_png(
            report_path, fx["cu"], fx["snap"],
            supplemental=fx.get("supplemental", False))

        for sheet, png in pngs.items():
            key = _safe(sheet)
            ref_path = fx_dir / f"{key}.png"
            if update or not ref_path.exists():
                ref_path.write_bytes(png)
                results.append({"fixture": label, "sheet": sheet,
                                "status": "captured"})
                continue
            frac, diff_png = _diff(ref_path.read_bytes(), png)
            status = "ok" if frac <= fail_fraction else "CHANGED"
            entry = {"fixture": label, "sheet": sheet, "status": status,
                     "changed_fraction": round(frac, 6)}
            if status == "CHANGED":
                ok = False
                cur_path = fx_dir / f"_current_{key}.png"
                cur_path.write_bytes(png)
                entry["current"] = str(cur_path)
                if diff_png:
                    diff_path = fx_dir / f"_diff_{key}.png"
                    diff_path.write_bytes(diff_png)
                    entry["diff"] = str(diff_path)
            results.append(entry)

    return {"ok": ok, "update": update, "results": results,
            "refs_dir": str(refs_root)}
