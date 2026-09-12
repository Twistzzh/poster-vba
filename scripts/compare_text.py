# -*- coding: utf-8 -*-
"""对照源图与渲染图：在给定文字带内测量暗像素 bbox，用于校准字号（评审一）。

用法:
    python compare_text.py <源图> <渲染图> [--band 名称:x1,y1,x2,y2 ...]

- `--band` 可重复。两图按同一坐标带分别量暗像素包围盒，宽高对比即可
  反推渲染字号相对源图是偏大还是偏小（宽度比 ≈ 字号比）。
- 不传 `--band` 时自动模式：按图高切成若干水平带逐带对比（快速全图扫）。
- 前提：两张图同宽高比（渲染图由 render_preview.py 按同画布比例产出）。

依赖: Pillow、numpy（隔离 venv 已预装）。
"""
import argparse
import os
import sys

import numpy as np
from PIL import Image


def bbox(img, box):
    x1, y1, x2, y2 = box
    lum = np.asarray(img.convert("RGB")).astype(np.float32).mean(axis=2)[y1:y2, x1:x2]
    bg = np.percentile(lum, 88)
    m = lum < bg - 55
    ys, xs = np.where(m)
    if len(xs) == 0:
        return None
    return (int(xs.max() - xs.min() + 1), int(ys.max() - ys.min() + 1),
            int(xs.min() + x1), int(xs.max() + x1))


def parse_band(s):
    name, coords = s.split(":", 1)
    x1, y1, x2, y2 = (int(v) for v in coords.split(","))
    return name, x1, y1, x2, y2


def main():
    ap = argparse.ArgumentParser(description="poster-vba: 文字带字号校准（与源图对照校准字号，评审一用）")
    ap.add_argument("src", help="源图路径")
    ap.add_argument("ren", help="渲染图路径")
    ap.add_argument("--band", action="append", default=[],
                    help="名称:x1,y1,x2,y2（可重复）")
    ap.add_argument("--band-h", type=int, default=80,
                    help="自动模式的带高（px，默认 80）")
    args = ap.parse_args()
    for p in (args.src, args.ren):
        if not os.path.isfile(p):
            print("找不到文件: %s" % p)
            return 2

    src = Image.open(args.src)
    ren = Image.open(args.ren)
    if args.band:
        bands = [parse_band(b) for b in args.band]
    else:
        h = min(src.height, ren.height)
        bands = [("auto y=%d" % y, 0, y, min(src.width, ren.width),
                  min(y + args.band_h, h))
                 for y in range(0, h - 10, args.band_h)]

    print("%-16s %-24s %-24s" % ("band", "source w,h", "render w,h"))
    big, small = [], []
    for name, x1, y1, x2, y2 in bands:
        a = bbox(src, (x1, y1, x2, y2))
        b = bbox(ren, (x1, y1, x2, y2))
        print("%-16s %-24s %-24s" % (name, a, b))
        if a and b and a[0] > 4 and b[0] > 4:
            r = b[0] / a[0]
            if r > 1.15:
                big.append((name, r))
            elif r < 0.85:
                small.append((name, r))
    if big:
        print("\n渲染偏大 (>1.15x): %s"
              % ", ".join("%s %.2fx" % (n, r) for n, r in big))
    if small:
        print("渲染偏小 (<0.85x): %s"
              % ", ".join("%s %.2fx" % (n, r) for n, r in small))
    if not big and not small:
        print("\n各带宽度比均在 0.85~1.15x 内，字号与源图基本一致。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
