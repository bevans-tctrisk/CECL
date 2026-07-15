"""Empty-dartboard base frame for the Barbecue Sauce launch video.

Redraws the print's dartboard (no darts, no text) as clean vectors using
geometry measured off the source art: center (2250, 2104), rim radius 1500,
bull 88 / bull ring 205 / triple band 735-855 / double band 1260-1375,
6px warm-gray hairlines. Black wedges carry red bands, cream wedges green,
with a black wedge centered at 12 and 6 o'clock, matching the print.

Outputs:
  found5_dartboard_clean_frame_4500x5400.png  (print-canvas layout)
  found5_dartboard_clean_frame_9x16.png       (2160x3840, video-ready)
"""

import math
from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).parent

S = 2  # supersample factor
CREAM = (237, 228, 207)
BLACK = (34, 35, 33)
RED = (177, 59, 53)
GREEN = (47, 107, 79)
HAIR = (142, 138, 126)
RIM = (34, 35, 33)
R_BULL, R_BULLRING = 88, 205
R_TRI_IN, R_TRI_OUT = 735, 855
R_DBL_IN, R_DBL_OUT = 1260, 1375
R_RIM = 1500
HW = 6  # hairline half-width at print scale


def main():
    sz = (2 * R_RIM + 8) * S
    im = Image.new("RGB", (sz, sz), (0, 0, 0))
    d = ImageDraw.Draw(im)
    c = sz // 2

    def disc(r, fill, a0=None, a1=None):
        box = (c - r * S, c - r * S, c + r * S, c + r * S)
        if a0 is None:
            d.ellipse(box, fill=fill)
        else:
            d.pieslice(box, a0, a1, fill=fill)

    def ring_line(r):
        d.ellipse((c - r * S - HW, c - r * S - HW, c + r * S + HW, c + r * S + HW),
                  outline=HAIR, width=2 * HW)

    disc(R_RIM, RIM)
    # Wedge boundaries at 81 + 18k degrees (PIL: clockwise from 3 o'clock,
    # y-down) put a black wedge dead-center at the bottom, as in the print.
    for i in range(20):
        a0 = 81 + 18 * i
        black_wedge = i % 2 == 0
        wedge = BLACK if black_wedge else CREAM
        band = RED if black_wedge else GREEN
        disc(R_DBL_OUT, band, a0, a0 + 18)
        disc(R_DBL_IN, wedge, a0, a0 + 18)
        disc(R_TRI_OUT, band, a0, a0 + 18)
        disc(R_TRI_IN, wedge, a0, a0 + 18)
    for i in range(20):
        ang = math.radians(81 + 18 * i)
        d.line((c + R_BULLRING * S * math.cos(ang), c + R_BULLRING * S * math.sin(ang),
                c + R_DBL_OUT * S * math.cos(ang), c + R_DBL_OUT * S * math.sin(ang)),
               fill=HAIR, width=2 * HW)
    for r in (R_TRI_IN, R_TRI_OUT, R_DBL_IN, R_DBL_OUT):
        ring_line(r)
    disc(R_BULLRING, GREEN)
    ring_line(R_BULLRING)
    disc(R_BULL, RED)
    ring_line(R_BULL)

    board = im.resize((sz // S, sz // S), Image.LANCZOS)

    cx, cy = 2250, 2104
    full = Image.new("RGB", (4500, 5400), (0, 0, 0))
    full.paste(board, (cx - board.width // 2, cy - board.height // 2))
    full.save(HERE / "found5_dartboard_clean_frame_4500x5400.png")

    v = Image.new("RGB", (2160, 3840), (0, 0, 0))
    b2 = board.resize((1560, 1560), Image.LANCZOS)
    v.paste(b2, ((2160 - 1560) // 2, (3840 - 1560) // 2 - 200))
    v.save(HERE / "found5_dartboard_clean_frame_9x16.png")
    print("done")


if __name__ == "__main__":
    main()
