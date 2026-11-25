"""Generate cross-platform WordCombine app icons (PNG, multi-res ICO, macOS ICNS)."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFont


def _draw_icon(size: int = 512) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Rounded Microsoft-Word blue squircle background
    margin = int(size * 0.06)
    radius = int(size * 0.18)
    draw.rounded_rectangle(
        [margin, margin, size - margin, size - margin],
        radius=radius,
        fill=(27, 94, 194, 255),
    )
    # Subtle lighter inner highlight band at top
    draw.rounded_rectangle(
        [margin + 6, margin + 6, size - margin - 6, size - margin - 6],
        radius=max(8, radius - 6),
        outline=(79, 146, 238, 180),
        width=max(2, size // 64),
    )

    # Back document sheet (slightly left-up)
    doc_w = int(size * 0.44)
    doc_h = int(size * 0.54)
    bx0 = int(size * 0.19)
    by0 = int(size * 0.17)
    draw.rounded_rectangle(
        [bx0, by0, bx0 + doc_w, by0 + doc_h],
        radius=int(size * 0.04),
        fill=(219, 234, 254, 255),
        outline=(147, 197, 253, 255),
        width=max(2, size // 128),
    )

    # Front document sheet (merged output, crisp white)
    fx0 = int(size * 0.36)
    fy0 = int(size * 0.28)
    draw.rounded_rectangle(
        [fx0, fy0, fx0 + doc_w, fy0 + doc_h],
        radius=int(size * 0.04),
        fill=(255, 255, 255, 255),
        outline=(30, 64, 175, 255),
        width=max(2, size // 96),
    )

    # Horizontal text lines on front document
    line_x0 = fx0 + int(doc_w * 0.16)
    line_x1 = fx0 + int(doc_w * 0.84)
    line_h = max(4, int(size * 0.028))
    for i, frac in enumerate((0.28, 0.44, 0.60, 0.76)):
        ly = fy0 + int(doc_h * frac)
        end_x = line_x1 if i != 3 else fx0 + int(doc_w * 0.62)
        color = (27, 94, 194, 255) if i == 0 else (100, 116, 139, 220)
        draw.rounded_rectangle([line_x0, ly, end_x, ly + line_h], radius=line_h // 2, fill=color)

    # Small green/blue merge badge on bottom-left
    badge_r = int(size * 0.14)
    cx = int(size * 0.29)
    cy = int(size * 0.71)
    draw.ellipse(
        [cx - badge_r, cy - badge_r, cx + badge_r, cy + badge_r],
        fill=(16, 185, 129, 255),
        outline=(255, 255, 255, 255),
        width=max(3, size // 64),
    )
    # Plus / combine symbol inside badge
    arm = int(badge_r * 0.52)
    thick = max(4, int(size * 0.036))
    draw.rounded_rectangle([cx - arm, cy - thick // 2, cx + arm, cy + thick // 2], radius=thick // 2, fill=(255, 255, 255, 255))
    draw.rounded_rectangle([cx - thick // 2, cy - arm, cx + thick // 2, cy + arm], radius=thick // 2, fill=(255, 255, 255, 255))

    return img


def main() -> None:
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    assets_dir = os.path.join(repo_root, "assets")
    os.makedirs(assets_dir, exist_ok=True)

    master = _draw_icon(512)
    png_path = os.path.join(assets_dir, "icon.png")
    master.resize((256, 256), Image.Resampling.LANCZOS).save(png_path, format="PNG")

    ico_path = os.path.join(assets_dir, "icon.ico")
    master.save(
        ico_path,
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )

    if sys.platform == "darwin" and shutil.which("iconutil"):
        iconset_dir = os.path.join(assets_dir, "WordCombine.iconset")
        os.makedirs(iconset_dir, exist_ok=True)
        try:
            for sz in (16, 32, 128, 256, 512):
                master.resize((sz, sz), Image.Resampling.LANCZOS).save(
                    os.path.join(iconset_dir, f"icon_{sz}x{sz}.png")
                )
                if sz <= 256:
                    dbl = sz * 2
                    master.resize((dbl, dbl), Image.Resampling.LANCZOS).save(
                        os.path.join(iconset_dir, f"icon_{sz}x{sz}@2x.png")
                    )
            icns_path = os.path.join(assets_dir, "icon.icns")
            subprocess.run(["iconutil", "-c", "icns", iconset_dir, "-o", icns_path], check=True)
        finally:
            shutil.rmtree(iconset_dir, ignore_errors=True)

    print("Generated icons in:", assets_dir)


if __name__ == "__main__":
    main()
