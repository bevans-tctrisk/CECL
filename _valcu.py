"""One-shot seed-driven validation for a WARM-based CU (read-mostly).

Builds a seed from the prior WARM quarter, runs the current quarter seed-driven,
and compares the recomputed pooled allowance to the current WARM's own pooled
allowance. Temporarily enables ``warm_seed_driven`` in the CU's yaml and restores
it afterward. Does NOT ship or commit anything.

Usage: python _valcu.py <short_name>
"""
import os, re, glob, sys, shutil, warnings
from calendar import monthrange
warnings.filterwarnings("ignore")

WS = os.environ["CECL_WORKSPACE_ROOT"]


def _snaps(cfg):
    fld = (cfg.get("credit_pull", {}) or {}).get("fallback_report_folder", "")
    months = set()
    for f in glob.glob(os.path.join(fld, "*.xlsx")):
        b = os.path.basename(f)
        if "WARM" not in b.upper() or b.startswith("~$") or b.upper().startswith("DNU"):
            continue
        m = re.search(r"(\d{4})-\s?(\d{2})", b)
        if m:
            months.add((int(m.group(1)), int(m.group(2))))
    ordered = sorted(months)
    if len(ordered) < 2:
        return None, None
    def iso(ym):
        y, mo = ym
        return f"{y:04d}-{mo:02d}-{monthrange(y, mo)[1]:02d}"
    return iso(ordered[-2]), iso(ordered[-1])  # prior, current


def main(short):
    from cecl_ui.services import config_service
    import generate_report as gr, warm_seed
    cfg = config_service.load_client_config(WS, short)
    prior, cur = _snaps(cfg)
    if not cur:
        print(f"{short}: could not resolve two WARM quarters"); return
    target = gr.load_impaired_data(cfg, cur).get("pooled_total_allowance")
    warm_seed.build_seed(short, prior)

    y = os.path.join(WS, "client_configs", f"{short}.yaml")
    bak = y + ".bak_val"
    shutil.copyfile(y, bak)
    import tempfile
    tmp_rpt = tempfile.mkdtemp(prefix=f"valcu_{short}_")
    _rpt_orig = gr.RPT_DIR
    try:
        with open(y, "a", encoding="utf-8") as fh:
            fh.write("\nwarm_seed_driven: true\n")
        import report_tct
        _o = report_tct.compute_acl_environmental
        cap = {}
        def _c(df, grades, config, hist, snap, *a, **k):
            r = _o(df, grades, config, hist, snap, *a, **k); cap["r"] = r; return r
        report_tct.compute_acl_environmental = _c
        # Redirect report output to a temp dir so validation never clobbers the
        # real Reports folder (a skipped render would otherwise leave 0-byte PDFs).
        gr.RPT_DIR = tmp_rpt
        import cecl_report_web.assembly as _asm
        _asm.render_report_pdf_from_data = lambda *a, **k: b""
        from cecl_ui.services import pipeline_service
        pipeline_service.run_reports(short, cur, ["tct_pdf"])
        report_tct.compute_acl_environmental = _o
    finally:
        gr.RPT_DIR = _rpt_orig
        shutil.rmtree(tmp_rpt, ignore_errors=True)
        shutil.copyfile(bak, y)
        os.remove(bak)

    s = cap.get("r", {}).get("acl_summary", {})
    p = s.get("pooled_total_allow") or 0
    drift = (p / target - 1) * 100 if target else 0
    print(f"RESULT {short}: prior={prior} cur={cur}")
    print(f"  seed-driven pooled={p:,.2f}  WARM target={target:,.2f}  drift={p-target:+,.2f} ({drift:+.1f}%)")
    print(f"  seed-driven spec={s.get('total_spec_allow',0):,.2f}  TAN={s.get('total_allow_needed',0):,.2f}  acl={s.get('acl_balance',0):,.2f}")
    print(f"  CONFIG restored (warm_seed_driven removed): {'warm_seed_driven' not in open(y,encoding='utf-8').read()}")


if __name__ == "__main__":
    main(sys.argv[1])
