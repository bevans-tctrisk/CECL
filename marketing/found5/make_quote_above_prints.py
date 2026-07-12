"""Found 5 Dollars darts print: quote above the board, 'Barbecue Sauce' below.

The punchline stays under the board; the Ted Lasso quote sits as a smaller
setup line in the clear band above the dart flights (which start at y=500).
Everything is set in Bevan. The source art is natively transparent, so the
alpha channel is preserved directly; the black-background version is a
composite for mockups.

Variants:
  quote_be_curious   - "Be Curious" above
  quote_full         - "Be curious, not judgmental" above
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
SRC = HERE / "found5_barbecue_sauce_print_source_4500x5400.png"
FONT = HERE / "fonts" / "Bevan-Regular.ttf"

CREAM = (237, 228, 207, 255)
CENTER_X = 2250
OLD_TITLE_BAND = (0, 3960, 4500, 4430)
TITLE_CY = 4193          # Barbecue Sauce center, as in the original print
TOP_CY = 235             # quote line center in the 0-500 band above the darts


def size_for_width(text, target_width):
    dummy = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    probe = ImageFont.truetype(str(FONT), 100)
    b = dummy.textbbox((0, 0), text, font=probe)
    size = int(100 * target_width / (b[2] - b[0]))
    while True:
        font = ImageFont.truetype(str(FONT), size)
        b = dummy.textbbox((0, 0), text, font=font)
        w = b[2] - b[0]
        if w > target_width:
            size -= 2
        elif w < target_width - 30:
            size += 2
        else:
            return font


def draw_centered(draw, text, font, cy):
    b = draw.textbbox((0, 0), text, font=font)
    w, h = b[2] - b[0], b[3] - b[1]
    draw.text((CENTER_X - w // 2 - b[0], cy - h // 2 - b[1]), text, font=font, fill=CREAM)


def build(quote, quote_width, tag):
    im = Image.open(SRC).convert("RGBA")
    d = ImageDraw.Draw(im)

    # Replace the old title with Bevan 'Barbecue Sauce' (band cleared to alpha 0).
    d.rectangle(OLD_TITLE_BAND, fill=(0, 0, 0, 0))
    draw_centered(d, "Barbecue Sauce", size_for_width("Barbecue Sauce", 3379), TITLE_CY)

    # Setup line above the board.
    draw_centered(d, quote, size_for_width(quote, quote_width), TOP_CY)

    out_t = HERE / f"found5_{tag}_print_BEVAN_transparent.png"
    im.save(out_t)
    print("wrote", out_t)

    black = Image.new("RGB", im.size, (0, 0, 0))
    black.paste(im, (0, 0), im)
    out_b = HERE / f"found5_{tag}_print_BEVAN_4500x5400.png"
    black.save(out_b)
    print("wrote", out_b)


def main():
    build("Be Curious", 1900, "quote_be_curious")
    build("Be curious, not judgmental", 3800, "quote_full")


if __name__ == "__main__":
    main()
