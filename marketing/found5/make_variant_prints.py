"""Title variants of the Found 5 Dollars darts print, set in Bevan.

Same board art and branding as the 'Barbecue Sauce' print; only the title
band changes. Generates black-background and transparent versions of:

  BE_CURIOUS   - "Be Curious" one line, same type size as Barbecue Sauce
  CURIOUS_FULL - "Be curious," / "not judgmental" stacked two lines
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
SRC = HERE / "found5_barbecue_sauce_print_source_4500x5400.png"
FONT = HERE / "fonts" / "Bevan-Regular.ttf"

CREAM = (237, 228, 207)
CENTER_X = 2250
CLEAR_BAND = (0, 3700, 4500, 4560)  # everything between board and icon
GAP_CENTER_Y = 4110                 # midpoint of board-bottom..icon-top


def size_for_width(text, target_width):
    """Font size at which `text` renders ~target_width wide."""
    probe = ImageFont.truetype(str(FONT), 100)
    dummy = ImageDraw.Draw(Image.new("RGB", (1, 1)))
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
            return size


def draw_centered(draw, text, font, cy):
    b = draw.textbbox((0, 0), text, font=font)
    w, h = b[2] - b[0], b[3] - b[1]
    draw.text((CENTER_X - w // 2 - b[0], cy - h // 2 - b[1]), text, font=font, fill=CREAM)
    return h


def save_variants(im, tag):
    out = HERE / f"found5_{tag}_print_BEVAN_4500x5400.png"
    im.convert("RGB").save(out)
    print("wrote", out)
    t = im.copy()
    px = t.load()
    for j in range(t.height):
        for i in range(t.width):
            r, g, b, a = px[i, j]
            if r + g + b <= 45:
                px[i, j] = (0, 0, 0, 0)
    out_t = HERE / f"found5_{tag}_print_BEVAN_transparent.png"
    t.save(out_t)
    print("wrote", out_t)


def fresh_canvas():
    im = Image.open(SRC).convert("RGBA")
    d = ImageDraw.Draw(im)
    d.rectangle(CLEAR_BAND, fill=(0, 0, 0, 255))
    return im, d


def main():
    # Type size the Barbecue Sauce title used, so 'Be Curious' matches it.
    base_size = size_for_width("Barbecue Sauce", 3379)

    # Variant 1: "Be Curious" - one line, original type size and baseline area.
    im, d = fresh_canvas()
    draw_centered(d, "Be Curious", ImageFont.truetype(str(FONT), base_size), 4193)
    save_variants(im, "be_curious")

    # Variant 2: full quote, two stacked lines centered in the gap.
    lines = ["Be curious,", "not judgmental"]
    size = size_for_width(lines[1], 3100)
    font = ImageFont.truetype(str(FONT), size)
    dummy = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    heights = [dummy.textbbox((0, 0), t, font=font)[3] - dummy.textbbox((0, 0), t, font=font)[1]
               for t in lines]
    gap = int(0.42 * max(heights))
    total = sum(heights) + gap
    im, d = fresh_canvas()
    cy = GAP_CENTER_Y - total // 2 + heights[0] // 2
    draw_centered(d, lines[0], font, cy)
    cy += heights[0] // 2 + gap + heights[1] // 2
    draw_centered(d, lines[1], font, cy)
    save_variants(im, "be_curious_not_judgmental")


if __name__ == "__main__":
    main()
