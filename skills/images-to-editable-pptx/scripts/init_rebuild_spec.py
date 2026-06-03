from __future__ import annotations

import argparse
import json
import os
import re
from pathlib import Path
from typing import Any

from PIL import Image


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


def natural_sort_key(path: Path) -> list[Any]:
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", path.name)
    ]


def find_images(images_dir: Path) -> list[Path]:
    return sorted(
        [
            path
            for path in images_dir.iterdir()
            if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
        ],
        key=natural_sort_key,
    )


def sample_background(image_path: Path) -> str:
    with Image.open(image_path).convert("RGB") as image:
        width, height = image.size
        points = [
            (0, 0),
            (max(width - 1, 0), 0),
            (0, max(height - 1, 0)),
            (max(width - 1, 0), max(height - 1, 0)),
        ]
        colors = [image.getpixel(point) for point in points]
    channels = [sorted(color[channel] for color in colors) for channel in range(3)]
    median = tuple(channel[len(channel) // 2] for channel in channels)
    return "#{:02X}{:02X}{:02X}".format(*median)


def create_spec(images_dir: Path, output_path: Path) -> dict[str, Any]:
    images = find_images(images_dir)
    if not images:
        raise ValueError(f"No slide images found in {images_dir}")

    slides = []
    for index, image_path in enumerate(images, start=1):
        relative_path = os.path.relpath(image_path, output_path.parent).replace("\\", "/")
        slides.append(
            {
                "name": f"Slide {index}",
                "source_image": relative_path,
                "background": sample_background(image_path),
                "elements": [],
            }
        )

    return {
        "title": images_dir.name,
        "slide_size": {"width": 13.333, "height": 7.5},
        "theme": {
            "font_face": "Microsoft YaHei",
            "text_color": "#0F172A",
            "accent_color": "#0D9488",
        },
        "slides": slides,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create an editable PPT reconstruction spec skeleton."
    )
    parser.add_argument("--images-dir", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()

    spec = create_spec(args.images_dir, args.out)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Created {args.out} with {len(spec['slides'])} slide references")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
