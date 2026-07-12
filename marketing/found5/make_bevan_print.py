"""Re-set the 'Barbecue Sauce' title on the Found 5 Dollars darts print in Bevan.

Keeps the dartboard art, app icon, and wordmarks untouched; only the title
band is cleared and re-rendered with the real Bevan typeface (OFL, Google
Fonts) in the print's cream ink.

Outputs:
  found5_barbecue_sauce_print_BEVAN_4500x5400.png       (black background)
  found5_barbecue_sauce_print_BEVAN_transparent.png     (for dark-garment DTG)
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
SRC = HERE / "found5_barbecue_sauce_print_source_4500x5400.png"
FONT = HERE / "fonts" / "Bevan-Regular.ttf"

CREAM = (237, 228, 207)
TITLE = "Barbecue Sauce"

# Measured off the source print: old title occupied x 561-3940, y 4036-4351.
TARGET_WIDTH = 3379
CENTER_X = 2250
CENTER_Y = 4193
CLEAR_BAND = (0, 3960, 4500, 4430)  # full-width band around the old title


def fit_font(draw, target_width):
    size = 100
    font = ImageFont.truetype(str(FONT), size)
    w = draw.textbbox((0, 0), TITLE, font=font)[2] - draw.textbbox((0, 0), TITLE, font=font)[0]
    size = int(size * target_width / w)
    while True:
        font = ImageFont.truetype(str(FONT), size)
        box = draw.textbbox((0, 0), TITLE, font=font)
        w = box[2] - box[0]
        if w > target_width:
            size -= 2
        elif w < target_width - 30:
            size += 2
        else:
            return font, box


def main():
    im = Image.open(SRC).convert("RGBA")
    draw = ImageDraw.Draw(im)

    draw.rectangle(CLEAR_BAND, fill=(0, 0, 0, 255))

    font, box = fit_font(draw, TARGET_WIDTH)
    w, h = box[2] - box[0], box[3] - box[1]
    x = CENTER_X - w // 2 - box[0]
    y = CENTER_Y - h // 2 - box[1]
    draw.text((x, y), TITLE, font=font, fill=CREAM)

    out = HERE / "found5_barbecue_sauce_print_BEVAN_4500x5400.png"
    im.convert("RGB").save(out)
    print("wrote", out)

    # Transparent version: knock out the black background for dark garments.
    px = im.load()
    for j in range(im.height):
        for i in range(im.width):
            r, g, b, a = px[i, j]
            if r + g + b <= 45:
                px[i, j] = (0, 0, 0, 0)
    out_t = HERE / "found5_barbecue_sauce_print_BEVAN_transparent.png"
    im.save(out_t)
    print("wrote", out_t)


if __name__ == "__main__":
    main()
