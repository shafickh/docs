#!/usr/bin/env python3
"""
Build the Vitacare Rondebosch signature lockup: the Vitacare logo with
"Responsible Pharmacist / Shafick Hassan (M.Pharm)" rendered into the same image.

Usage:
    python3 build_lockup.py LOGO.png --layout beside --out lockup.png
    python3 build_lockup.py LOGO.png --layout under  --out lockup.png

Layouts:
    beside  full-width banner: name block on the left, Vitacare logo on the right
    under   logo on top, name block right-aligned beneath it

The image is rendered at 2x and the HTML displays it at 1x, so it stays sharp on
high-resolution screens. Everything is drawn on transparency, so it sits on
whatever background the mail client uses.
"""
import argparse
from PIL import Image, ImageDraw, ImageFont

# Carlito and Caladea are metric-compatible open substitutes for Calibri and
# Cambria. They are what gets baked into the image; on screen the difference
# from the real Microsoft fonts is not perceptible at these sizes.
SANS_BOLD = "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf"
SERIF_BOLD = "/usr/share/fonts/truetype/crosextra/Caladea-Bold.ttf"

ORANGE = (192, 66, 26, 255)
INK = (26, 26, 26, 255)

SCALE = 2           # render at 2x, display at 1x
LOGO_W = 320        # logo display width in px
GAP = 16            # gap between name block and logo


def _fonts(s):
    return (
        ImageFont.truetype(SERIF_BOLD, int(19 * s)),   # "Responsible Pharmacist"
        ImageFont.truetype(SANS_BOLD, int(25 * s)),    # "Shafick Hassan"
        ImageFont.truetype(SANS_BOLD, int(18 * s)),    # "(M.Pharm)"
    )


def _name_block(s):
    """Render the title + name onto its own transparent tile."""
    f_title, f_name, f_qual = _fonts(s)
    tmp = ImageDraw.Draw(Image.new("RGBA", (1, 1)))

    title = "Responsible Pharmacist"
    name = "Shafick Hassan"
    qual = " (M.Pharm)"

    w_title = tmp.textlength(title, font=f_title)
    w_name = tmp.textlength(name, font=f_name)
    w_qual = tmp.textlength(qual, font=f_qual)

    width = int(max(w_title, w_name + w_qual)) + 2
    line1_h = int(26 * s)
    height = line1_h + int(34 * s)

    tile = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    d = ImageDraw.Draw(tile)
    d.text((0, 0), title, font=f_title, fill=ORANGE)
    # baseline-align the qualification with the name
    d.text((0, line1_h), name, font=f_name, fill=INK)
    d.text((w_name, line1_h + int(6 * s)), qual, font=f_qual, fill=INK)
    return tile


def build(logo_path, layout, out_path, total_width=720):
    s = SCALE
    logo = Image.open(logo_path).convert("RGBA")
    logo_w = LOGO_W * s
    logo_h = round(logo.height * logo_w / logo.width)
    logo = logo.resize((logo_w, logo_h), Image.LANCZOS)

    block = _name_block(s)

    if layout == "beside":
        height = max(block.height, logo_h)
        # stretch to the full signature width so the banner lines up with the
        # orange strap underneath it
        width = max(total_width * s, block.width + GAP * s + logo_w)
        canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        canvas.paste(block, (0, height - block.height), block)      # bottom-aligned
        canvas.paste(logo, (width - logo_w, height - logo_h), logo)
    elif layout == "under":
        width = max(block.width, logo_w)
        height = logo_h + int(6 * s) + block.height
        canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        canvas.paste(logo, (width - logo_w, 0), logo)
        canvas.paste(block, (width - block.width, logo_h + int(6 * s)), block)
    else:
        raise SystemExit(f"unknown layout: {layout}")

    canvas.save(out_path)
    print(f"{out_path}  {canvas.width}x{canvas.height}px  (display at "
          f"{canvas.width // s}x{canvas.height // s})")
    return canvas.width // s, canvas.height // s


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("logo")
    ap.add_argument("--layout", default="beside", choices=["beside", "under"])
    ap.add_argument("--out", default="lockup.png")
    ap.add_argument("--width", type=int, default=720,
                    help="display width of the beside banner, in px")
    a = ap.parse_args()
    build(a.logo, a.layout, a.out, a.width)
