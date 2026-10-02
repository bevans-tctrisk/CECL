import re


def build(snap_prefix):
    year = snap_prefix[:4]
    m = int(snap_prefix[5:7])
    names = {
        1: r"jan(?:uary)?", 2: r"feb(?:ruary)?", 3: r"mar(?:ch)?",
        4: r"apr(?:il)?", 5: r"may", 6: r"jun(?:e)?",
        7: r"jul(?:y)?", 8: r"aug(?:ust)?",
        9: r"sep(?:t(?:ember)?)?", 10: r"oct(?:ober)?",
        11: r"nov(?:ember)?", 12: r"dec(?:ember)?",
    }
    mn = names[m]
    alts = [
        rf"{year}[-_.\s]*0?{m}(?!\d)",
        rf"(?<!\d)0?{m}[-_.\s]*{year}",
        rf"{mn}[-_.\s]*{year}",
        rf"{year}[-_.\s]*{mn}",
    ]
    date_rx = re.compile("|".join(alts), re.IGNORECASE)
    imp_rx = re.compile(r"Impaired[\s_\-]+Loans", re.IGNORECASE)
    mm = f"{m:02d}"
    strict = re.compile(
        rf"^{year}\s*[-_]?\s*{mm}.*Impaired[\s_-]+Loans.*\.xlsx$",
        re.IGNORECASE,
    )
    return strict, imp_rx, date_rx


def test(fname, snap):
    strict, imp_rx, date_rx = build(snap)
    if "WARM" in fname.upper():
        strict_hit = False
        loose_hit = False
        note = "  (WARM skip)"
    else:
        strict_hit = bool(strict.match(fname))
        loose_hit = bool(imp_rx.search(fname) and date_rx.search(fname))
        note = ""
    print(f"  {fname!r:70} strict={strict_hit} loose={loose_hit}{note}")


for snap in ("2026-06", "2025-12", "2026-03"):
    print(f"=== snap={snap} ===")
    for fname in (
        "Impaired Loans - 6-2026.xlsx",
        "Impaired Loans - 06-2026.xlsx",
        "Impaired Loans - June 2026.xlsx",
        "Impaired Loans - Jun_2026.xlsx",
        "Impaired Loans 2026-06.xlsx",
        "2026-06 Impaired Loans - TCP CU.xlsx",
        "2025-12_CECL_Migration_Impaired_Loans_-_TCP_CU_65384.xlsx",
        "2026-03 CECL Migration Impaired Loans - TCP CU 65384.xlsx",
        "CECL-Migration-WARM - TCP CU.xlsx",
        "Impaired Loans - 1-2026.xlsx",  # January — must NOT match snap=2026-06
    ):
        test(fname, snap)
