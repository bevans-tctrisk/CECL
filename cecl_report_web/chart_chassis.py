"""Reusable chart chassis: scales, axes, ticks, formatting, palette, labels.

The machinery every archetype in the CECL Migration deck needs, factored out
once so the seven renderers are configuration rather than geometry. Nothing
here knows about openpyxl or workbooks -- it consumes the plain spec dict
described below, so the same renderers work against a parsed .xlsx today and
against the ``ReportData`` model after the Step 1a extraction.

Contents
--------
* ``PALETTE`` / ``THEME``   -- validated colour tokens (see NOTE ON COLOUR)
* ``text_width`` etc.       -- real Calibri advance widths, so label placement
                               is *measured* rather than guessed
* ``LinearScale``/``BandScale`` -- domain -> pixel mapping, sign-aware
* ``nice_ticks``            -- 1/2/2.5/5/10 axis tick generation
* ``fmt_*`` / ``axis_formatter`` -- number formatting with ONE unit chosen for
                               the whole axis, not per tick
* ``Frame``                 -- box model; margins can be measured from content
* ``y_axis`` / ``x_category_axis`` / ``legend`` -- chrome components
* ``place_bar_label`` / ``resolve_1d_collisions`` -- label placement
* ``render_diverging_stacked_bar`` (A2) and ``render_clustered_column`` (A5)

THE SPEC DICT
-------------
Every renderer takes one dict of this shape::

    {
      "kind":        "diverging_stacked_bar" | "clustered_column" | ...,
      "title":       str | None,
      "subtitle":    str | None,
      "categories":  [str, ...],
      "series": [
        {
          "name":      str,
          "values":    [float, ...],      # parallel to categories
          "color":     "#rrggbb" | None,  # None -> palette slot by index
          "filled":    bool,              # False -> outline-only marks
          "direction": +1 | -1,           # diverging charts only; see below
        }, ...
      ],
      "value_format": "pct1" | "pct0" | "currency" | "number" | <excel numfmt>,
      "axis_title":   str | None,
      "width":        int,
      "height":       int,
      "options":      {...}               # per-archetype knobs
    }

NOTE ON THE SIGN CONVENTION (the A2 bug this module exists to kill)
-------------------------------------------------------------------
In the Excel source the "Deteriorated" series is stored NEGATIVE, and Excel
draws a diverging tornado as a side effect of stacking mixed signs. The old
prototype took ``abs(val)`` and accumulated into a single running total, so
Improved and Deteriorated ADDED instead of OPPOSING -- a materially wrong
chart that looks plausible.

``split_stack`` below is the fix: it keeps two independent accumulators, one
growing positive from zero and one growing negative from zero, so segments on
opposite sides of the baseline can never contaminate each other.

The renderer must ALSO not depend on the feed's sign convention. Each series
carries an optional ``direction`` (+1/-1); the effective plotted value is
``value * direction`` after taking ``abs()`` when a direction is declared. So
a future ``ReportData`` that supplies Deteriorated as a positive magnitude
renders identically to today's negative-signed workbook cells, by declaring
``direction: -1``. See ``effective_values``.

NOTE ON COLOUR
--------------
The Vizo brand hexes used by the Excel charts (teal ``0D4D5E``, maroon
``873A3A``/``3D1A1A``, olive ``829901``, amber ``FFC000``) do NOT pass a
categorical palette validation: the teal is below the chroma floor (0.066,
reads gray) and outside the lightness band, and olive/amber are
indistinguishable under deuteranopia (dE 1.5).

The steps below are the same four hues re-stepped to pass. Validated with the
dataviz skill's ``validate_palette.js`` on surface #ffffff-ish, order
teal / maroon / amber / olive:

    lightness band  PASS   chroma floor    PASS
    CVD separation  PASS (worst adjacent olive<->amber dE 12.3 protan)
    normal-vision   PASS (worst adjacent dE 19.6)
    contrast        WARN  amber #E0A400 at 2.16:1 -> "relief" required, i.e.
                    amber-coded marks MUST carry a visible direct label or
                    appear in the table view. A6/A7 (the four-slice DQ pie
                    and charge-off bar) are the only archetypes that reach
                    the amber slot, and both label every slice, so the
                    relief obligation is met -- but re-run the validator if
                    that changes.

Teal/maroon is also a legitimate *diverging* pair for A2: cool vs warm poles
reading as opposite, with a neutral gray zero rule as the midpoint.
"""

from __future__ import annotations

import html as _html
import math
from typing import Any, Callable, Iterable, Sequence

# ── Colour tokens ────────────────────────────────────────────────────

#: Categorical slots, in assignment order. Never cycle past the end --
#: fold a 5th category into "Other" (none of the seven archetypes needs one).
PALETTE = [
    "#0E7E9E",  # 1 teal    (brand 0D4D5E, re-stepped for chroma/lightness)
    "#B4453F",  # 2 maroon  (brand 873A3A)
    "#E0A400",  # 3 amber   (brand FFC000) -- contrast WARN, needs labels
    "#6E8A00",  # 4 olive   (brand 829901)
]

#: Semantic assignment for the migration-direction charts. "Improved" is the
#: brand olive (829901 -> re-stepped PALETTE[3]) and "Deteriorated" the brand
#: maroon, matching the risk-change matrix, the DQ/CO pies and the Net Credit
#: Change doughnut so one hue means one thing across every page.
SEMANTIC = {
    "improved": PALETTE[3],
    "deteriorated": PALETTE[1],
    "unchanged": "#9A9A93",
}

THEME = {
    "surface": "#ffffff",
    "ink": "#1a1a18",        # primary text
    "ink_secondary": "#55554e",
    "ink_muted": "#8a8a82",  # axis tick labels
    "on_color": "#ffffff",   # text drawn on top of a saturated fill
    "grid": "#e6e6e1",       # hairline gridlines, one shade off surface
    "axis": "#c9c9c2",       # axis rule
    "zero": "#6f6f68",       # the diverging baseline -- deliberately stronger
}

FONT = "Calibri, Carlito, 'Segoe UI', sans-serif"


# ── Text measurement (real Calibri metrics) ──────────────────────────
# Advance widths in em, extracted from the bundled static/fonts/calibri.ttf
# (hmtx/cmap). Label placement in this module is genuinely measured; the old
# prototype used fixed offsets and a blind ``[:7]`` truncation, which clips
# pool names like "Consumer Indirect Auto-Used Auto".

CALIBRI_EM: dict[str, float] = {
    " ": 0.2261, "!": 0.3257, "\"": 0.4009, "#": 0.498, "$": 0.5068,
    "%": 0.7148, "&": 0.6821, "'": 0.2207, "(": 0.3032, ")": 0.3032,
    "*": 0.498, "+": 0.498, ",": 0.2495, "-": 0.3062, ".": 0.2524,
    "/": 0.3862, "0": 0.5068, "1": 0.5068, "2": 0.5068, "3": 0.5068,
    "4": 0.5068, "5": 0.5068, "6": 0.5068, "7": 0.5068, "8": 0.5068,
    "9": 0.5068, ":": 0.2676, ";": 0.2676, "<": 0.498, "=": 0.498,
    ">": 0.498, "?": 0.4429, "@": 0.8443, "A": 0.5786, "B": 0.5679,
    "C": 0.5513, "D": 0.6304, "E": 0.5165, "F": 0.4956, "G": 0.6416,
    "H": 0.6484, "I": 0.2861, "J": 0.3179, "K": 0.5674, "L": 0.4595,
    "M": 0.855, "N": 0.6885, "O": 0.687, "P": 0.5528, "Q": 0.6899,
    "R": 0.5757, "S": 0.4936, "T": 0.523, "U": 0.6675, "V": 0.5762,
    "W": 0.8896, "X": 0.5225, "Y": 0.4941, "Z": 0.4873, "[": 0.3032,
    "\\": 0.3862, "]": 0.3032, "^": 0.498, "_": 0.498, "`": 0.3335,
    "a": 0.4795, "b": 0.5166, "c": 0.4272, "d": 0.5166, "e": 0.4976,
    "f": 0.3057, "g": 0.4595, "h": 0.5171, "i": 0.2295, "j": 0.2295,
    "k": 0.4595, "l": 0.2295, "m": 0.7993, "n": 0.5171, "o": 0.5162,
    "p": 0.5166, "q": 0.5166, "r": 0.3535, "s": 0.3853, "t": 0.3364,
    "u": 0.5171, "v": 0.4429, "w": 0.7261, "x": 0.4331, "y": 0.4429,
    "z": 0.3838, "{": 0.3125, "|": 0.4595, "}": 0.3125, "~": 0.498,
    " ": 0.2261, "–": 0.498, "—": 0.9727, "−": 0.498,
}

_DEFAULT_EM = 0.52          # fallback for anything outside the table
_BOLD_FACTOR = 1.045        # calibrib is ~4.5% wider on average


def text_width(s: Any, size: float, *, bold: bool = False) -> float:
    """Measured advance width of ``s`` in px at ``size`` px, in Calibri."""
    if s is None:
        return 0.0
    total = 0.0
    for ch in str(s):
        total += CALIBRI_EM.get(ch, _DEFAULT_EM)
    w = total * float(size)
    return w * _BOLD_FACTOR if bold else w


def ellipsize(s: Any, size: float, max_w: float, *, bold: bool = False) -> str:
    """Truncate to fit ``max_w`` px, appending a real ellipsis if it had to."""
    s = "" if s is None else str(s)
    if text_width(s, size, bold=bold) <= max_w:
        return s
    ell = "…"
    ew = text_width(ell, size, bold=bold)
    out = ""
    w = 0.0
    for ch in s:
        cw = CALIBRI_EM.get(ch, _DEFAULT_EM) * size * (_BOLD_FACTOR if bold else 1.0)
        if w + cw + ew > max_w:
            break
        out += ch
        w += cw
    return (out.rstrip() + ell) if out else ell


# ── Scales ───────────────────────────────────────────────────────────

class LinearScale:
    """Map a numeric domain onto a pixel range.

    Sign-agnostic by construction: ``d0`` may be negative and ``r0`` may be
    greater than ``r1`` (a reversed / "maxMin" axis), so the same object
    serves an upward column axis, a rightward bar axis and a diverging
    horizontal axis whose zero sits mid-range.
    """

    __slots__ = ("d0", "d1", "r0", "r1", "_k")

    def __init__(self, d0: float, d1: float, r0: float, r1: float):
        self.d0, self.d1 = float(d0), float(d1)
        self.r0, self.r1 = float(r0), float(r1)
        span = self.d1 - self.d0
        self._k = 0.0 if span == 0 else (self.r1 - self.r0) / span

    def __call__(self, v: float) -> float:
        return self.r0 + (float(v) - self.d0) * self._k

    def length(self, dv: float) -> float:
        """Pixel length of a domain *delta* (always non-negative)."""
        return abs(float(dv) * self._k)

    @property
    def zero(self) -> float:
        """Pixel position of domain 0 -- the baseline every signed chart hangs off."""
        return self(0.0)


class BandScale:
    """Evenly spaced categorical bands (one per category)."""

    __slots__ = ("n", "r0", "r1", "pad_outer", "pad_inner", "_step")

    def __init__(self, n: int, r0: float, r1: float, *,
                 pad_outer: float = 0.08, pad_inner: float = 0.0):
        self.n = max(1, int(n))
        self.r0, self.r1 = float(r0), float(r1)
        self.pad_outer, self.pad_inner = pad_outer, pad_inner
        self._step = (self.r1 - self.r0) / self.n

    def start(self, i: int) -> float:
        return self.r0 + i * self._step

    def center(self, i: int) -> float:
        return self.r0 + (i + 0.5) * self._step

    @property
    def step(self) -> float:
        return self._step

    def band(self) -> float:
        """Usable width inside one band, after outer padding."""
        return abs(self._step) * (1.0 - 2.0 * self.pad_outer)


# ── Ticks ────────────────────────────────────────────────────────────

_NICE_MANTISSAS = (1.0, 2.0, 2.5, 5.0, 10.0)


def nice_step(raw: float) -> float:
    """Round a raw step up to the next 1 / 2 / 2.5 / 5 / 10 x 10^k."""
    if raw <= 0 or not math.isfinite(raw):
        return 1.0
    exp = math.floor(math.log10(raw))
    frac = raw / (10.0 ** exp)
    for m in _NICE_MANTISSAS:
        if frac <= m + 1e-9:
            return m * (10.0 ** exp)
    return 10.0 ** (exp + 1)


def nice_ticks(lo: float, hi: float, *, target: int = 5,
               symmetric: bool = False,
               include_zero: bool = True) -> tuple[list[float], float, float]:
    """Return ``(ticks, domain_lo, domain_hi)`` over a nice rounded domain.

    ``symmetric`` forces ``-M .. +M`` so both halves of a diverging chart
    share one scale -- without it a 4% bar on the left could be drawn longer
    than a 4% bar on the right, which is the classic tornado-chart lie.
    """
    lo, hi = float(lo), float(hi)
    if not math.isfinite(lo) or not math.isfinite(hi):
        lo, hi = 0.0, 1.0
    if include_zero and not symmetric:
        lo, hi = min(lo, 0.0), max(hi, 0.0)
    if symmetric:
        m = max(abs(lo), abs(hi))
        lo, hi = -m, m
    if hi == lo:
        hi = lo + 1.0 if lo == 0 else lo + abs(lo) * 0.5
        if symmetric:
            lo = -hi

    target = max(2, int(target))
    span = hi - lo
    raw = span / target if span > 0 else 1.0
    exp = math.floor(math.log10(raw)) if raw > 0 else 0

    # Score every candidate step rather than blindly rounding the raw step
    # UP: rounding up alone turns a 15.2% maximum into a 20% domain with
    # ticks only every 10%. Candidates one decade either side let the engine
    # trade tick count against how much empty domain it adds.
    best: tuple[float, float, float, float] | None = None
    for e in (exp - 1, exp, exp + 1):
        for m in _NICE_MANTISSAS[:-1]:
            step_c = m * (10.0 ** e)
            if step_c <= 0:
                continue
            a = math.floor(lo / step_c) * step_c
            b = math.ceil(hi / step_c) * step_c
            if symmetric:
                mm = max(abs(a), abs(b))
                a, b = -mm, mm
            count = int(round((b - a) / step_c))
            if count < 2 or count > 12:
                continue
            overshoot = ((b - a) / span - 1.0) if span > 0 else 0.0
            score = abs(count + 1 - target) + overshoot * 3.0
            if best is None or score < best[0]:
                best = (score, step_c, a, b)
    if best is None:
        step = nice_step(raw)
        d0 = math.floor(lo / step) * step
        d1 = math.ceil(hi / step) * step
        if symmetric:
            m = max(abs(d0), abs(d1))
            d0, d1 = -m, m
    else:
        _, step, d0, d1 = best

    ticks: list[float] = []
    n = int(round((d1 - d0) / step))
    for i in range(n + 1):
        v = d0 + i * step
        # kill -0.0 and float dust so labels and the zero test behave
        if abs(v) < step * 1e-9:
            v = 0.0
        ticks.append(round(v, 12))
    return ticks, d0, d1


# ── Number formatting ────────────────────────────────────────────────

def fmt_pct(v: float, decimals: int = 1, *, signed: bool = False) -> str:
    try:
        n = float(v) * 100.0
    except (TypeError, ValueError):
        return ""
    s = f"{abs(n):.{decimals}f}%"
    if signed and abs(n) >= 10.0 ** -decimals / 2:
        s = ("−" if n < 0 else "+") + s
    elif n < 0 and not signed:
        s = "−" + s          # true minus sign, not a hyphen
    return s


_UNITS = ((1e9, "B"), (1e6, "M"), (1e3, "K"), (1.0, ""))


def currency_unit(magnitude: float) -> tuple[float, str]:
    """Pick ONE divisor+suffix for a whole axis from its largest value.

    Choosing per-tick produces axes reading "$500K, $1.0M, $1.5M" -- mixed
    units on one scale, which is a readability bug. Callers pick once and
    format every tick and label through it.
    """
    m = abs(float(magnitude or 0))
    for div, suf in _UNITS:
        if m >= div:
            return div, suf
    return 1.0, ""


def fmt_currency(v: float, *, div: float = 1.0, suffix: str = "",
                 decimals: int | None = None) -> str:
    """Format one value against a pre-chosen unit (see :func:`currency_unit`)."""
    try:
        n = float(v)
    except (TypeError, ValueError):
        return ""
    if n == 0:
        return "$0"          # never "$0.0M" -- zero has no unit
    scaled = n / div
    if decimals is None:
        decimals = 0 if (div == 1.0 or abs(scaled) >= 100) else 1
    s = f"${abs(scaled):,.{decimals}f}{suffix}"
    return ("(" + s + ")") if n < 0 else s


def decimals_for_step(step: float, *, cap: int = 3) -> int:
    """Fewest decimals that render ``step`` without rounding it away.

    A 2.5-percentage-point tick step formatted with 0 decimals prints
    "2%, 5%, 8%, 10%" -- Python's round-half-to-even turns 2.5 into "2" and
    7.5 into "8", so an evenly spaced axis reads as an uneven one. Deriving
    the precision from the step instead of guessing kills that for every
    percent and currency axis in the deck.

    >>> decimals_for_step(5.0), decimals_for_step(2.5), decimals_for_step(0.25)
    (0, 1, 2)
    """
    step = abs(float(step))
    if step <= 0:
        return 0
    for d in range(cap + 1):
        if abs(round(step, d) - step) < step * 1e-9:
            return d
    return cap


def tick_formatter(kind: str | None,
                   ticks: Sequence[float]) -> Callable[[float], str]:
    """Formatter whose precision comes from the tick STEP, not the data.

    An axis stepping by 5% should print "5%", "10%" -- not "5.0%", "10.0%" --
    while the per-bar labels on the same chart still carry a decimal. Getting
    this from the step is what keeps the two consistent without a second
    hand-set format.
    """
    kind = (kind or "").strip()
    step = min((abs(b - a) for a, b in zip(ticks, ticks[1:])), default=0.0)
    if step <= 0:
        return axis_formatter(kind, ticks)
    if kind.startswith("pct"):
        return lambda v, d=decimals_for_step(step * 100.0): fmt_pct(v, d)
    if kind in ("currency", "$"):
        peak = max((abs(float(t)) for t in ticks), default=0.0)
        div, suf = currency_unit(peak)
        dec = decimals_for_step(step / div)
        return lambda v, _d=dec: fmt_currency(v, div=div, suffix=suf, decimals=_d)
    return axis_formatter(kind, ticks)


def axis_formatter(kind: str | None, values: Iterable[float]) -> Callable[[float], str]:
    """Return a single formatter closed over the whole axis's unit choice.

    ``kind`` accepts the chassis names ("pct0".."pct2", "currency", "number")
    and falls through to :func:`cecl_report_web.format.excel_format` for a
    raw Excel number-format mask, so specs read straight off a workbook still
    work.
    """
    kind = (kind or "").strip()
    vals = [abs(float(v)) for v in values if v is not None]
    peak = max(vals) if vals else 0.0

    if kind.startswith("pct"):
        dec = int(kind[3:]) if kind[3:].isdigit() else 1
        return lambda v: fmt_pct(v, dec)
    if kind in ("currency", "$"):
        div, suf = currency_unit(peak)
        return lambda v: fmt_currency(v, div=div, suffix=suf)
    if kind == "number":
        return lambda v: f"{float(v):,.0f}"
    if kind:
        from .format import excel_format
        return lambda v: excel_format(v, kind)
    return lambda v: f"{float(v):,.2f}"


# ── Frame / box model ────────────────────────────────────────────────

class Frame:
    """Chart box model: outer size plus margins, yielding the plot rect.

    Margins are meant to be *computed* from measured content (widest tick
    label, longest category name) rather than hard-coded, which is how the
    prototype's fixed 340x240 viewBox ends up clipping pool names.
    """

    __slots__ = ("w", "h", "top", "right", "bottom", "left")

    def __init__(self, w: float, h: float, *, top: float = 8, right: float = 8,
                 bottom: float = 8, left: float = 8):
        self.w, self.h = float(w), float(h)
        self.top, self.right = float(top), float(right)
        self.bottom, self.left = float(bottom), float(left)

    @property
    def x0(self) -> float:
        return self.left

    @property
    def x1(self) -> float:
        return self.w - self.right

    @property
    def y0(self) -> float:
        return self.top

    @property
    def y1(self) -> float:
        return self.h - self.bottom

    @property
    def pw(self) -> float:
        return max(1.0, self.x1 - self.x0)

    @property
    def ph(self) -> float:
        return max(1.0, self.y1 - self.y0)


# ── SVG primitives ───────────────────────────────────────────────────

def esc(s: Any) -> str:
    return _html.escape("" if s is None else str(s))


def _f(v: float) -> str:
    """Compact fixed-point, so golden-SVG diffs stay stable."""
    return f"{v:.2f}".rstrip("0").rstrip(".") or "0"


def svg_text(x: float, y: float, s: Any, *, size: float = 9,
             fill: str | None = None, anchor: str = "start",
             weight: int = 400, opacity: float | None = None,
             rotate: float | None = None,
             extra: str = "") -> str:
    fill = fill or THEME["ink"]
    a = "" if anchor == "start" else f' text-anchor="{anchor}"'
    w = "" if weight == 400 else f' font-weight="{weight}"'
    o = "" if opacity is None else f' opacity="{_f(opacity)}"'
    r = "" if rotate is None else f' transform="rotate({_f(rotate)} {_f(x)} {_f(y)})"'
    return (f'<text x="{_f(x)}" y="{_f(y)}" font-size="{_f(size)}" '
            f'fill="{fill}"{a}{w}{o}{r}{extra}>{esc(s)}</text>')


def svg_rect(x: float, y: float, w: float, h: float, *, fill: str = "none",
             stroke: str | None = None, stroke_width: float = 1.4,
             rx: float = 0, extra: str = "") -> str:
    s = "" if not stroke else f' stroke="{stroke}" stroke-width="{_f(stroke_width)}"'
    r = "" if not rx else f' rx="{_f(rx)}"'
    return (f'<rect x="{_f(x)}" y="{_f(y)}" width="{_f(max(0.0, w))}" '
            f'height="{_f(max(0.0, h))}" fill="{fill}"{s}{r}{extra}/>')


def svg_line(x1: float, y1: float, x2: float, y2: float, *,
             stroke: str, width: float = 1.0, extra: str = "") -> str:
    return (f'<line x1="{_f(x1)}" y1="{_f(y1)}" x2="{_f(x2)}" y2="{_f(y2)}" '
            f'stroke="{stroke}" stroke-width="{_f(width)}"{extra}/>')


def svg_open(w: float, h: float, *, cls: str = "chart") -> str:
    return (f'<svg class="{cls}" xmlns="http://www.w3.org/2000/svg" '
            f'viewBox="0 0 {_f(w)} {_f(h)}" width="{_f(w)}" height="{_f(h)}" '
            f'font-family="{FONT}" role="img">')


def svg_close() -> str:
    return "</svg>"


# ── Chrome components ────────────────────────────────────────────────

TITLE_SIZE = 12
SUBTITLE_SIZE = 9
TICK_SIZE = 8.5
LABEL_SIZE = 8
LEGEND_SIZE = 9
GAP = 2.0          # the surface gap between adjacent / stacked fills


def header(frame: Frame, title: str | None, subtitle: str | None,
           *, x: float | None = None, anchor: str = "start") -> tuple[str, float]:
    """Draw title/subtitle. Returns (svg, y of the next free line)."""
    if not title and not subtitle:
        return "", 0.0
    x = frame.x0 if x is None else x
    parts = []
    y = TITLE_SIZE + 2
    if title:
        parts.append(svg_text(x, y, title, size=TITLE_SIZE, weight=600,
                              fill=THEME["ink"], anchor=anchor))
        y += SUBTITLE_SIZE + 4
    if subtitle:
        parts.append(svg_text(x, y, subtitle, size=SUBTITLE_SIZE,
                              fill=THEME["ink_secondary"], anchor=anchor))
        y += 4
    return "".join(parts), y


def legend(items: Sequence[tuple[str, str]], x: float, y: float, *,
           filled: bool = True, size: float = LEGEND_SIZE,
           gap: float = 16.0) -> str:
    """Horizontal legend. Always present for >=2 series (identity is never
    carried by colour alone)."""
    out: list[str] = []
    cx = x
    sw = size * 0.9
    for label, color in items:
        if filled:
            out.append(svg_rect(cx, y - sw + 1.5, sw, sw, fill=color, rx=1.5))
        else:
            out.append(svg_rect(cx + 0.7, y - sw + 2.2, sw - 1.4, sw - 1.4,
                                fill="none", stroke=color, stroke_width=1.6,
                                rx=1.5))
        out.append(svg_text(cx + sw + 4, y, label, size=size,
                            fill=THEME["ink_secondary"]))
        cx += sw + 4 + text_width(label, size) + gap
    return "".join(out)


def legend_width(items: Sequence[tuple[str, str]], *, size: float = LEGEND_SIZE,
                 gap: float = 16.0) -> float:
    sw = size * 0.9
    return sum(sw + 4 + text_width(lb, size) + gap for lb, _ in items) - gap


def y_axis(frame: Frame, scale: LinearScale, ticks: Sequence[float],
           fmt: Callable[[float], str], *, grid: bool = True,
           tick_len: float = 3.0, axis_title: str | None = None) -> str:
    """Left value axis: hairline gridlines, short ticks, right-aligned labels.

    This is the capability the Excel originals delete outright on A1, A2, A3,
    A5 and A7 (``<c:delete val="1"/>``) -- A5 in particular plots millions of
    dollars with no scale, no gridline and no number.
    """
    parts: list[str] = []
    for t in ticks:
        y = scale(t)
        if grid:
            parts.append(svg_line(frame.x0, y, frame.x1, y,
                                  stroke=THEME["grid"], width=1))
        parts.append(svg_line(frame.x0 - tick_len, y, frame.x0, y,
                              stroke=THEME["axis"], width=1))
        parts.append(svg_text(frame.x0 - tick_len - 3, y + TICK_SIZE * 0.35,
                              fmt(t), size=TICK_SIZE, anchor="end",
                              fill=THEME["ink_muted"],
                              extra=' style="font-variant-numeric:tabular-nums"'))
    parts.append(svg_line(frame.x0, frame.y0, frame.x0, frame.y1,
                          stroke=THEME["axis"], width=1))
    if axis_title:
        cx, cy = frame.x0 - max_tick_width(ticks, fmt) - 14, (frame.y0 + frame.y1) / 2
        parts.append(svg_text(cx, cy, axis_title, size=TICK_SIZE, anchor="middle",
                              fill=THEME["ink_muted"], rotate=-90))
    return "".join(parts)


def max_tick_width(ticks: Sequence[float], fmt: Callable[[float], str],
                   *, size: float = TICK_SIZE) -> float:
    """Widest formatted tick -- feed this into the Frame's left margin."""
    return max([text_width(fmt(t), size) for t in ticks] + [0.0])


def x_category_axis(frame: Frame, band: BandScale, cats: Sequence[str], *,
                    size: float = LABEL_SIZE, y: float | None = None,
                    rule: bool = True) -> str:
    """Bottom category axis, measured: labels that do not fit are rotated
    45 degrees before they are allowed to be truncated."""
    y = frame.y1 + size + 4 if y is None else y
    parts: list[str] = []
    if rule:
        parts.append(svg_line(frame.x0, frame.y1, frame.x1, frame.y1,
                              stroke=THEME["axis"], width=1))
    avail = abs(band.step) - 3
    too_wide = any(text_width(c, size) > avail for c in cats)
    for i, c in enumerate(cats):
        cx = band.center(i)
        if too_wide:
            parts.append(svg_text(cx, y, ellipsize(c, size, 84), size=size,
                                  anchor="end", fill=THEME["ink_secondary"],
                                  rotate=-45))
        else:
            parts.append(svg_text(cx, y, c, size=size, anchor="middle",
                                  fill=THEME["ink_secondary"]))
    return "".join(parts)


def category_axis_height(cats: Sequence[str], band_step: float,
                         *, size: float = LABEL_SIZE) -> float:
    """How much bottom margin :func:`x_category_axis` will actually need."""
    avail = abs(band_step) - 3
    if any(text_width(c, size) > avail for c in cats):
        longest = max([min(text_width(c, size), 84) for c in cats] + [0.0])
        return longest * 0.7071 + size + 6      # rotated 45 degrees
    return size + 8


# ── Signed stacking + label placement ────────────────────────────────

def effective_values(series: dict, n: int) -> list[float]:
    """Values with the spec's ``direction`` applied, padded to ``n``.

    When a series declares a direction the magnitude is taken from the data
    and the SIDE from the spec, so the renderer is immune to whether the feed
    stores Deteriorated as -0.0345 (today's workbook) or +0.0345 (a sane data
    model). Without a direction the value's own sign is authoritative.
    """
    vals = list(series.get("values") or [])
    vals = (vals + [0.0] * n)[:n]
    d = series.get("direction")
    out: list[float] = []
    for v in vals:
        try:
            f = float(v)
        except (TypeError, ValueError):
            f = 0.0
        out.append(abs(f) * (1 if d > 0 else -1) if d else f)
    return out


def split_stack(row: Sequence[float]) -> list[tuple[float, float, float]]:
    """Lay one category's segments out around zero. THE signed-baseline core.

    Returns ``[(from, to, value), ...]`` in domain units, where positives
    accumulate upward/rightward from 0 and negatives accumulate
    downward/leftward from 0 through an INDEPENDENT accumulator.

    The bug this replaces::

        acc = 0
        for v in row:               # prototype
            w = abs(v) / vmax * pw  # sign thrown away
            draw(acc, acc + w)
            acc += w                # +6.5% and -3.5% ADD to a 10% bar

    Here they oppose, which is the chart Excel actually draws::

        >>> split_stack([0.0647, -0.0345])
        [(0.0, 0.0647, 0.0647), (0.0, -0.0345, -0.0345)]
        >>> split_stack([3.0, 2.0, -1.0, -4.0])       # multi-segment sides
        [(0.0, 3.0, 3.0), (3.0, 5.0, 2.0), (0.0, -1.0, -1.0), (-1.0, -5.0, -4.0)]
    """
    pos = 0.0
    neg = 0.0
    out: list[tuple[float, float, float]] = []
    for v in row:
        v = float(v or 0.0)
        if v >= 0:
            out.append((pos, pos + v, v))
            pos += v
        else:
            out.append((neg, neg + v, v))
            neg += v
    return out


def stack_extent(rows: Sequence[Sequence[float]]) -> tuple[float, float]:
    """(most negative, most positive) reach across every stacked row."""
    lo = hi = 0.0
    for row in rows:
        p = sum(v for v in row if v > 0)
        n = sum(v for v in row if v < 0)
        hi = max(hi, p)
        lo = min(lo, n)
    return lo, hi


def place_bar_label(a: float, b: float, text: str, *, size: float = LABEL_SIZE,
                    pad: float = 4.0, outward: int = 1,
                    on_color: str | None = None,
                    outside_fill: str | None = None) -> dict | None:
    """Decide inside-vs-outside for one bar-end label, by measurement.

    ``a`` is the baseline end of the mark, ``b`` the far end, in pixels.
    ``outward`` is +1 when the mark grows toward increasing pixels.

    Excel's version of this is a magic ratio (``T_BAR = 0.30`` / ``T_COL =
    0.22``: "a bar shorter than 30% of the longest gets a black outside
    label"), which mis-fires whenever the label text length varies -- "0.0%"
    and "(1,234,567)" are treated identically. Here the text is measured
    against the space it has, so the rule is right by construction. SVG also
    has no equivalent of Excel's ban on ``outEnd`` for stacked bars, so the
    fallback is a genuine outside label rather than Excel's ``inEnd`` fudge.

    Returns ``{"x", "anchor", "fill", "inside"}`` or None if there is no room
    either way.
    """
    span = abs(b - a)
    tw = text_width(text, size)
    if span >= tw + 2 * pad:
        return {"x": b - outward * pad, "inside": True,
                "anchor": "end" if outward > 0 else "start",
                "fill": on_color or THEME["on_color"]}
    return {"x": b + outward * pad, "inside": False,
            "anchor": "start" if outward > 0 else "end",
            "fill": outside_fill or THEME["ink_secondary"]}


def resolve_1d_collisions(positions: Sequence[float], min_gap: float,
                          *, lo: float | None = None,
                          hi: float | None = None) -> list[float]:
    """Nudge a sorted-by-input list of 1-D label positions apart.

    A single forward pass then a backward clamp -- enough for axis ticks and
    pie leader labels, which is where every archetype needs it. Order is
    preserved, so a label never crosses its neighbour.
    """
    if not positions:
        return []
    idx = sorted(range(len(positions)), key=lambda i: positions[i])
    out = list(positions)
    prev = None
    for i in idx:
        v = out[i]
        if prev is not None and v - prev < min_gap:
            v = prev + min_gap
        if hi is not None and v > hi:
            v = hi
        out[i] = v
        prev = v
    prev = None
    for i in reversed(idx):
        v = out[i]
        if prev is not None and prev - v < min_gap:
            v = prev - min_gap
        if lo is not None and v < lo:
            v = lo
        out[i] = v
        prev = v
    return out


def _color(series: Sequence[dict], i: int) -> str:
    c = series[i].get("color")
    if c:
        return c
    name = str(series[i].get("name") or "").strip().lower()
    if name in SEMANTIC:
        return SEMANTIC[name]
    return PALETTE[i % len(PALETTE)]


def _mark_attrs(s: dict, color: str) -> dict:
    if s.get("filled", True):
        return {"fill": color, "stroke": None}
    return {"fill": "none", "stroke": color}


# ── A2: diverging stacked bar ────────────────────────────────────────

def render_diverging_stacked_bar(spec: dict) -> str:
    """Horizontal stacked bar around a signed zero baseline (archetype A2).

    What this fixes relative to both the Excel original and the prototype:

    * segments oppose across zero instead of adding (``split_stack``);
    * one SYMMETRIC scale, so a 4% left bar is exactly as long as a 4% right
      bar -- Excel had no value axis at all and could not be checked;
    * a real value axis with ticks and hairline gridlines, labelled in
      magnitude (direction is carried by side + legend, as in a population
      pyramid);
    * category labels on the LEFT in reading order. Excel put them on the
      right via ``catAx axPos="r"`` AND reversed the value axis so positives
      grew leftward; neither is ported;
    * measured label placement instead of the ``T_BAR = 0.30`` heuristic and
      its 16 hand-written per-point overrides;
    * an explicit empty state for the all-zero pools (3 of 11 here), which
      Excel draws as an invisible nothing.
    """
    cats = [str(c) for c in (spec.get("categories") or [])]
    series = list(spec.get("series") or [])
    n = len(cats)
    opts = spec.get("options") or {}
    w = float(spec.get("width") or 620)

    rows = [[effective_values(s, n)[i] for s in series] for i in range(n)]
    fmt = axis_formatter(spec.get("value_format") or "pct1",
                         [v for r in rows for v in r])

    lo, hi = stack_extent(rows)
    ticks, d0, d1 = nice_ticks(lo, hi, target=opts.get("ticks", 8),
                               symmetric=True)
    tfmt = tick_formatter(spec.get("value_format") or "pct1", ticks)

    # Measured margins: category column on the left, room on the right for an
    # outside label that does not fit inside its segment.
    cat_size = float(opts.get("cat_size", 8.5))
    cat_w = min(max([text_width(c, cat_size) for c in cats] + [0.0]) + 10,
                float(opts.get("cat_max", 190)))
    max_lbl = max([text_width(fmt(v), LABEL_SIZE) for r in rows for v in r] + [0.0])
    side_pad = max_lbl + 8

    row_h = float(opts.get("row_height", 22))
    legend_items = [(str(s.get("name") or f"Series {i+1}"), _color(series, i))
                    for i, s in enumerate(series)]
    head_svg, head_y = header(Frame(w, 0), spec.get("title"), spec.get("subtitle"),
                              x=w / 2, anchor="middle")
    top = max(head_y, 4) + LEGEND_SIZE + 10
    bottom = TICK_SIZE + 16
    h = float(spec.get("height") or (top + n * row_h + bottom))

    frame = Frame(w, h, top=top, left=cat_w + side_pad, right=side_pad + 4,
                  bottom=bottom)
    x = LinearScale(d0, d1, frame.x0, frame.x1)
    band = BandScale(n, frame.y0, frame.y1)
    bar_h = min(band.band(), float(opts.get("bar_max", 16)))

    p: list[str] = [svg_open(w, h),
                    svg_rect(0, 0, w, h, fill=THEME["surface"]), head_svg]

    # legend, top-left of the plot
    p.append(legend(legend_items, frame.x0, top - 6,
                    filled=series[0].get("filled", True) if series else True))

    # gridlines + value ticks (magnitudes; sign lives in the legend)
    for t in ticks:
        tx = x(t)
        is_zero = abs(t) < 1e-12
        p.append(svg_line(tx, frame.y0 - 2, tx, frame.y1,
                          stroke=THEME["zero"] if is_zero else THEME["grid"],
                          width=1.2 if is_zero else 1))
        p.append(svg_text(tx, frame.y1 + TICK_SIZE + 4,
                          "0" if is_zero else tfmt(abs(t)),
                          size=TICK_SIZE, anchor="middle",
                          fill=THEME["ink_muted"],
                          extra=' style="font-variant-numeric:tabular-nums"'))

    zero_x = x.zero
    for i, cat in enumerate(cats):
        cy = band.center(i)
        y = cy - bar_h / 2
        # category label, left column, measured + ellipsized (never a [:7] cut)
        p.append(svg_text(frame.x0 - side_pad - 8, cy + cat_size * 0.35,
                          ellipsize(cat, cat_size, cat_w), size=cat_size,
                          anchor="end", fill=THEME["ink"]))

        segs = split_stack(rows[i])
        if all(abs(v) < 1e-12 for _, _, v in segs):
            # honest empty state -- Excel renders these three pools as nothing
            p.append(svg_text(zero_x + 6, cy + LABEL_SIZE * 0.35,
                              "no migration", size=LABEL_SIZE - 0.5,
                              fill=THEME["ink_muted"], opacity=0.9))
            continue

        for si, (a, b, v) in enumerate(segs):
            if abs(v) < 1e-12:
                continue
            xa, xb = x(a), x(b)
            outward = 1 if xb >= xa else -1
            # 2px surface gap at the zero end so opposing segments never touch
            x_lo, x_hi = (xa, xb) if xb >= xa else (xb, xa)
            if outward > 0:
                x_lo += GAP / 2
            else:
                x_hi -= GAP / 2
            attrs = _mark_attrs(series[si], _color(series, si))
            p.append(svg_rect(x_lo, y, x_hi - x_lo, bar_h,
                              fill=attrs["fill"], stroke=attrs["stroke"],
                              stroke_width=1.6, rx=2))
            txt = fmt(abs(v))
            on_col = (THEME["on_color"] if series[si].get("filled", True)
                      else THEME["ink"])
            pl = place_bar_label(xa, xb, txt, outward=outward, on_color=on_col)
            p.append(svg_text(pl["x"], cy + LABEL_SIZE * 0.36, txt,
                              size=LABEL_SIZE, anchor=pl["anchor"],
                              fill=pl["fill"], weight=600 if pl["inside"] else 400,
                              extra=' style="font-variant-numeric:tabular-nums"'))

    if spec.get("axis_title"):
        p.append(svg_text((frame.x0 + frame.x1) / 2, h - 3, spec["axis_title"],
                          size=TICK_SIZE, anchor="middle", fill=THEME["ink_muted"]))
    p.append(svg_close())
    return "".join(p)


# ── A5: clustered column with a real value axis ──────────────────────

def render_clustered_column(spec: dict) -> str:
    """Grouped vertical columns with a proper value scale (archetype A5).

    The Excel original deletes the value axis (``<c:delete val="1"/>``), has
    no gridlines and no data labels, so nine charts plot millions of dollars
    with no number anywhere on them. This adds:

    * nice 1/2/2.5/5/10 ticks over a zero-anchored domain;
    * hairline gridlines behind the marks;
    * currency labels with ONE unit chosen for the whole axis ($0.0M ..
      $4.0M), not a per-tick unit;
    * a left margin measured from the widest tick label;
    * optional per-column value labels that are drawn only when they
      measurably fit, so the axis carries the rest;
    * a legend (2 series -> identity is never colour-alone).

    ``options["outline"]`` keeps Excel's hollow 3pt strokes for comparison;
    the default is filled, which is markedly more legible at report size.
    """
    cats = [str(c) for c in (spec.get("categories") or [])]
    series = list(spec.get("series") or [])
    n = len(cats)
    opts = spec.get("options") or {}
    outline = bool(opts.get("outline"))
    w = float(spec.get("width") or 560)
    h = float(spec.get("height") or 300)

    cols = [effective_values(s, n) for s in series]
    flat = [v for c in cols for v in c]
    lo, hi = min(flat + [0.0]), max(flat + [0.0])
    ticks, d0, d1 = nice_ticks(lo, hi, target=opts.get("ticks", 5))
    vfmt = spec.get("value_format") or "currency"
    # axis ticks take their precision from the STEP ("$1M"); the per-column
    # labels keep the data's own precision ("$3.9M")
    tfmt = tick_formatter(vfmt, ticks)
    fmt = axis_formatter(vfmt, ticks + flat)

    legend_items = [(str(s.get("name") or f"Series {i+1}"), _color(series, i))
                    for i, s in enumerate(series)]
    head_svg, head_y = header(Frame(w, 0), spec.get("title"), spec.get("subtitle"))
    top = max(head_y, 4) + LEGEND_SIZE + 12
    left = max_tick_width(ticks, tick_formatter(
        spec.get("value_format") or "currency", ticks)) + 12 + (14 if spec.get("axis_title") else 0)
    right = 10

    # Two-pass bottom margin: the category axis needs the band step, which
    # needs the plot width, which needs this margin -- so measure with a
    # provisional frame, then rebuild.
    prov = BandScale(n, left, w - right)
    bottom = category_axis_height(cats, prov.step) + 6
    frame = Frame(w, h, top=top, left=left, right=right, bottom=bottom)

    yscale = LinearScale(d0, d1, frame.y1, frame.y0)   # inverted: +y is up
    band = BandScale(n, frame.x0, frame.x1, pad_outer=0.10)
    ns = max(1, len(series))
    bar_w = max(2.0, (band.band() - GAP * (ns - 1)) / ns)

    p: list[str] = [svg_open(w, h),
                    svg_rect(0, 0, w, h, fill=THEME["surface"]), head_svg]
    p.append(legend(legend_items, frame.x0, top - 8, filled=not outline))
    p.append(y_axis(frame, yscale, ticks, tfmt,
                    axis_title=spec.get("axis_title")))

    base_y = yscale(0.0)
    label_size = LABEL_SIZE - 0.5
    for i in range(n):
        gx = band.center(i) - (bar_w * ns + GAP * (ns - 1)) / 2
        for si, s in enumerate(series):
            v = cols[si][i]
            bx = gx + si * (bar_w + GAP)
            vy = yscale(v)
            y_top, y_bot = min(vy, base_y), max(vy, base_y)
            color = _color(series, si)
            if abs(v) < 1e-12:
                # a visible zero tick beats an invisible bar: "Platinum
                # improved = $0" is information, not absence of data
                p.append(svg_line(bx, base_y, bx + bar_w, base_y,
                                  stroke=THEME["ink_muted"], width=1))
                continue
            if outline:
                p.append(svg_rect(bx + 0.8, y_top + 0.8, bar_w - 1.6,
                                  max(0.0, y_bot - y_top - 0.8), fill="none",
                                  stroke=color, stroke_width=1.6, rx=1.5))
            else:
                p.append(svg_rect(bx, y_top, bar_w, y_bot - y_top,
                                  fill=color, rx=2))
            if opts.get("value_labels", "auto") != "never":
                txt = fmt(v)
                # measured: only label when it fits the column's own width
                if text_width(txt, label_size) <= bar_w + GAP:
                    p.append(svg_text(bx + bar_w / 2, y_top - 3, txt,
                                      size=label_size, anchor="middle",
                                      fill=THEME["ink_secondary"],
                                      extra=' style="font-variant-numeric:tabular-nums"'))
    p.append(x_category_axis(frame, band, cats))
    p.append(svg_close())
    return "".join(p)


RENDERERS: dict[str, Callable[[dict], str]] = {
    "diverging_stacked_bar": render_diverging_stacked_bar,
    "clustered_column": render_clustered_column,
}


def render(spec: dict) -> str:
    """Dispatch a chassis spec to its archetype renderer."""
    kind = spec.get("kind")
    try:
        fn = RENDERERS[kind]
    except KeyError:
        raise ValueError(
            f"no chassis renderer for kind={kind!r}; "
            f"have {sorted(RENDERERS)}") from None
    return fn(spec)
