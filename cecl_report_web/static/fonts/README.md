# Bundled fonts for the browser report renderer

Headless Chromium does **not** use OS-installed fonts, so the report CSS
embeds fonts via `@font-face` (see `render.py`). The font files must be
present in this folder at render time.

## Local / workstation (current)
The renderer expects the Calibri family (matching the Excel reports):
`calibri.ttf`, `calibrib.ttf`, `calibrii.ttf`, `calibriz.ttf`. On Windows
these are copied from `C:\Windows\Fonts`. **Calibri is Microsoft-licensed
and is intentionally NOT committed to the repo** (see `.gitignore`).

## Shared server / redistribution (planned)
For a multi-user server, replace Calibri with **Carlito** — a
metric-compatible, freely redistributable substitute (same advance
widths, so line-breaks and pagination stay identical). Drop the Carlito
TTFs here and point the `_FONT_FACES` table in `render.py` at them
(keeping the CSS family name `Calibri`, or renaming consistently).
Carlito *can* be committed.
