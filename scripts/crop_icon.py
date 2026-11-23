import argparse
from pathlib import Path

from PIL import Image


def recover_alpha(image, background, foreground):
    delta = tuple(f - b for b, f in zip(background, foreground))
    scale = sum(d * d for d in delta)
    if not scale:
        raise ValueError("Foreground and background colors must differ")
    output = Image.new("RGBA", image.size)
    output.putdata([
        (*foreground, round(alpha * max(0, min(1, sum(
            (c - b) * d for c, b, d in zip(rgb, background, delta)
        ) / scale))))
        for *rgb, alpha in image.getdata()
    ])
    return output


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--box", nargs=4, type=int, required=True)
    parser.add_argument("--background")
    parser.add_argument("--foreground")
    parser.add_argument("--size", type=int)
    args = parser.parse_args()
    source = Image.open(args.source).convert("RGBA")
    x, y, width, height = args.box
    if min(x, y) < 0 or min(width, height) <= 0 or x + width > source.width or y + height > source.height:
        parser.error("Crop box is outside the image")
    if bool(args.background) != bool(args.foreground):
        parser.error("Specify both key colors")
    icon = source.crop((x, y, x + width, y + height))
    if args.background:
        colors = [bytes.fromhex(s.lstrip("#")) for s in (args.background, args.foreground)]
        if any(len(c) != 3 for c in colors):
            parser.error("Use six-digit RGB colors")
        icon = recover_alpha(icon, *colors)
    if args.size is not None:
        if not 1 <= args.size <= 4096:
            parser.error("Size must be 1–4096")
        icon.thumbnail((args.size, args.size), Image.Resampling.LANCZOS)
        canvas = Image.new("RGBA", (args.size, args.size))
        canvas.alpha_composite(icon, ((args.size-icon.width)//2, (args.size-icon.height)//2))
        icon = canvas
    args.output.parent.mkdir(parents=True, exist_ok=True)
    icon.save(args.output, format="PNG")


if __name__ == "__main__":
    main()
