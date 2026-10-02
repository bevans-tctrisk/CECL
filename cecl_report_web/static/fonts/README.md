# Bundled fonts for the browser report renderer

Headless Chromium does **not** use OS-installed fonts, so the report CSS
embeds fonts via `@font-face` (see `render.py`). The font files must be
present in this folder at render time.

## Local / workstation (current)
The renderer expects:
- **Calibri** family (TCT reports): `calibri.ttf`, `calibrib.ttf`,
  `calibrii.ttf`, `calibriz.ttf`.
- **Arial** family (Vizo body text, "Theme 2026" minor font): `arial.ttf`,
  `arialbd.ttf`, `ariali.ttf`, `arialbi.ttf`.
- **Montserrat** (Vizo headings, "Theme 2026" major font):
  `Montserrat-Variable.ttf`, `Montserrat-Italic-Variable.ttf` (SIL OFL,
  from google/fonts; committed, see `Montserrat-OFL.txt`).

On Windows the Calibri/Arial files are copied from `C:\Windows\Fonts`.
**They are Microsoft-licensed and intentionally NOT committed** (see
`.gitignore`).

## Shared server / redistribution (planned)
For a multi-user server, replace Calibri with **Carlito** — a
metric-compatible, freely redistributable substitute (same advance
widths, so line-breaks and pagination stay identical). Drop the Carlito
TTFs here and point the `_FONT_FACES` table in `render.py` at them
(keeping the CSS family name `Calibri`, or renaming consistently).
Carlito *can* be committed.
