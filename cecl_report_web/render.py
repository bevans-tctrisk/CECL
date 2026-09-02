"""Rendering layer: ReportModel/HTML -> PDF via headless Chromium.

Stateless helpers so the exact same code runs on a workstation now and a
shared Flask server later. Nothing here touches Excel or the filesystem
beyond reading the package's own templates/CSS.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape

_PKG_DIR = Path(__file__).resolve().parent
_TEMPLATES_DIR = _PKG_DIR / "templates"
_STATIC_DIR = _PKG_DIR / "static"
_FONTS_DIR = _STATIC_DIR / "fonts"

# Bundled font faces embedded as @font-face so headless Chromium renders
# identically on a workstation and a server (it does NOT pick up
# OS-installed fonts). NOTE: Calibri is Microsoft-licensed — fine on a
# Windows workstation, but for a redistributable multi-user server swap
# these TTFs for "Carlito" (metric-compatible, freely redistributable)
# and rename the family below; pagination stays identical.
_FONT_FACES = [
    # (file, weight, style)
    ("calibri.ttf", 400, "normal"),
    ("calibrib.ttf", 700, "normal"),
    ("calibrii.ttf", 400, "italic"),
    ("calibriz.ttf", 700, "italic"),
]


def _env() -> Environment:
    env = Environment(
        loader=FileSystemLoader(str(_TEMPLATES_DIR)),
        autoescape=select_autoescape(["html", "xml"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )
    from . import format as _fmt
    env.filters["acct0"] = _fmt.acct0
    env.filters["acct2"] = _fmt.acct2
    env.filters["pct0"] = _fmt.pct0
    env.filters["pct1"] = _fmt.pct1
    env.filters["pct2"] = _fmt.pct2
    env.filters["pct4"] = _fmt.pct4
    env.filters["mcell"] = _fmt.mcell
    env.filters["tcell"] = _fmt.tcell
    return env


def _embedded_font_css() -> str:
    """Build @font-face rules pointing at the bundled TTFs via file://.

    We render through a temp .html file (see :func:`render_pdf`) so the
    page origin is ``file://`` and Chromium will load these sibling
    font files. This is far more robust than multi-MB base64 data URIs
    (which Chromium's stylesheet parser silently dropped) and keeps the
    HTML small. On a server the fonts ship inside the package, so the
    absolute path resolves wherever it is deployed.
    """
    faces: list[str] = []
    for fname, weight, style in _FONT_FACES:
        fp = _FONTS_DIR / fname
        if not fp.exists():
            continue
        uri = fp.as_uri()  # file:///C:/.../calibri.ttf
        faces.append(
            "@font-face{"
            "font-family:'Calibri';"
            f"font-weight:{weight};font-style:{style};"
            f"src:url('{uri}') format('truetype');"
            "}"
        )
    return "".join(faces)


def load_css() -> str:
    """Return the print stylesheet with embedded fonts prepended.

    Inlining (fonts + CSS) keeps the HTML self-contained so Chromium
    needs no base URL or static-file server — important for
    headless/server portability.
    """
    css = _STATIC_DIR / "report.css"
    body = css.read_text(encoding="utf-8") if css.exists() else ""
    return _embedded_font_css() + "\n" + body


def render_html(template: str, **ctx: Any) -> str:
    """Render a template to a standalone HTML string (CSS inlined)."""
    ctx.setdefault("inline_css", load_css())
    return _env().get_template(template).render(**ctx)


def render_pdf(
    html: str,
    *,
    page_format: str = "Letter",
    landscape: bool = False,
    margin: str = "0.75in",
    print_background: bool = True,
) -> bytes:
    """Convert an HTML string to PDF bytes with headless Chromium.

    Uses Chromium's print-to-PDF, which honours ``@page`` size/margins,
    page-break CSS, web fonts, and inline SVG — the highest-fidelity
    option for "keep the formatting the same".
    """
    from playwright.sync_api import sync_playwright

    import tempfile

    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            page = browser.new_page()
            # Render via a temp file so the page origin is file:// and the
            # bundled @font-face TTFs (also file://) load same-origin.
            with tempfile.TemporaryDirectory() as td:
                fp = Path(td) / "report.html"
                fp.write_text(html, encoding="utf-8")
                page.goto(fp.as_uri(), wait_until="load")
                page.evaluate("document.fonts.ready")
                page.emulate_media(media="print")
                pdf_bytes = page.pdf(
                    format=page_format,
                    landscape=landscape,
                    print_background=print_background,
                    # NOTE: do NOT set prefer_css_page_size — it makes
                    # Chromium honor the CSS ``@page size`` (portrait) and
                    # silently ignore the ``landscape`` flag. Letting the
                    # pdf() format+landscape params drive page size keeps
                    # wide tabs (Risk Change, ACL Env, ...) landscape.
                    margin={
                        "top": margin,
                        "bottom": margin,
                        "left": margin,
                        "right": margin,
                    },
                )
        finally:
            browser.close()
    return pdf_bytes


def render_cover_pdf(cover, **pdf_opts: Any) -> bytes:
    """Convenience: render a :class:`CoverPage` straight to PDF bytes."""
    html = render_html("cover.html", cover=cover)
    return render_pdf(html, **pdf_opts)
