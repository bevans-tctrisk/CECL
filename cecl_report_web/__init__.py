"""Browser-rendered CECL report package.

Renders CECL Credit Migration reports as HTML (Jinja2) and converts them
to PDF via headless Chromium (Playwright). Designed to run identically on
a single workstation today and a shared Flask server later — the render
path is stateless and has no Excel/COM/pywin32 dependency.

Architecture (three layers, deliberately decoupled):

    compute  ->  ReportModel  ->  renderer
    (engine)     (this pkg's     (HTML/CSS -> Chromium -> PDF,
                  data contract)   or model -> xlsx for "extract page")

The ``ReportModel`` (see :mod:`cecl_report_web.model`) is the single
contract between the calculation engine and any renderer. Because every
page flows through the model, dumping a page back to Excel stays a small
adapter — that keeps the "PDF-only, but leave the Excel door open" goal
cheap.
"""

from .model import ReportModel, CoverPage  # noqa: F401
