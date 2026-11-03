#!/usr/bin/env python3
"""Pad images with transparent background to a target aspect ratio.

The input image is centered on a transparent canvas. Only the missing side
is expanded, so the original content is never distorted or cropped.
"""

from __future__ import annotations

import argparse
import math
from pathlib import Path

from PIL import Image


def parse_ratio(raw: str) -> float:
    if ":" in raw:
        left, right = raw.split(":", 1)
        return float(left) / float(right)
    return float(raw)


def pad_image(path: Path, target_ratio: float, overwrite: bool = True) -> bool:
    with Image.open(path) as src:
        src = src.convert("RGBA")
        width, height = src.size
        current_ratio = width / height if height else target_ratio
        if math.isclose(current_ratio, target_ratio, rel_tol=0.001, abs_tol=0.001):
            return False

        if current_ratio > target_ratio:
            new_width = width
            new_height = math.ceil(width / target_ratio)
        else:
            new_height = height
            new_width = math.ceil(height * target_ratio)

        canvas = Image.new("RGBA", (new_width, new_height), (0, 0, 0, 0))
        left = (new_width - width) // 2
        top = (new_height - height) // 2
        canvas.alpha_composite(src, (left, top))

        out_path = path if overwrite else path.with_name(f"{path.stem}_padded{path.suffix}")
        canvas.save(out_path)
        return True


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", help="Image files or directories to process.")
    parser.add_argument("--ratio", default="3:4", help="Target aspect ratio as W:H or decimal. Default: 3:4.")
    parser.add_argument("--inplace", action="store_true", help="Overwrite images in place.")
    args = parser.parse_args()

    target_ratio = parse_ratio(args.ratio)
    inputs: list[Path] = []
    for raw in args.paths:
        p = Path(raw)
        if p.is_dir():
            inputs.extend(sorted(x for x in p.glob("*.png") if x.is_file()))
        elif p.is_file():
            inputs.append(p)

    changed = 0
    for p in inputs:
        if pad_image(p, target_ratio, overwrite=args.inplace):
            changed += 1
            print(f"padded {p}")
        else:
            print(f"kept   {p}")

    print(f"done: {changed} padded, {len(inputs) - changed} unchanged")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
