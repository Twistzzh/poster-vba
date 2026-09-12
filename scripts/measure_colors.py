# -*- coding: utf-8 -*-
"""poster-vba 通用取色：整图主色 + 任意区域的底色/文字色/强调色。

用法:
    python measure_colors.py <img> [--top 12] [--box 名称:x1,y1,x2,y2 ...]

- 整图: 16 色自适应量化统计主色（含占比，排序输出），另给 bg_guess
  （边界 2px 环的中位色，用于判断展板底色）。
- 每个 --box 输出三个色:
    median  区域中位色 —— 面板/容器填充色
    dark4   最暗 4% 像素均值 —— 文字/描边色
    sat     最大饱和度簇均值 —— 强调色（标题、图标主色）

输出 JSON，直接对照写进 modPoster_Content*.bas 的调色板常量。
依赖: Pillow、numpy（隔离 venv 已预装）。
"""
import argparse
import json
import sys

import numpy as np
from PIL import Image


def dominant(img: Image.Image, top: int):
    q = img.convert("RGB").quantize(colors=16, method=Image.MEDIANCUT)
    pal = q.getpalette()
    counts = sorted(q.getcolors(maxcolors=16) or [], reverse=True)
    total = sum(c for c, _ in counts)
    out = []
    for cnt, idx in counts[:top]:
        r, g, b = pal[idx * 3:idx * 3 + 3]
        out.append({"rgb": [r, g, b], "share": round(cnt / total, 3)})
    return out


def box_colors(a: np.ndarray, box):
    x1, y1, x2, y2 = box
    sub = a[y1:y2, x1:x2].reshape(-1, 3).astype(np.float32)
    if len(sub) == 0:
        return None
    med = np.median(sub, axis=0)
    lum = sub.sum(axis=1)
    dark_sel = sub[lum <= np.percentile(lum, 4)]
    dark = dark_sel.mean(axis=0) if len(dark_sel) else med
    mx = sub.max(axis=1)
    mn = sub.min(axis=1)
    sat = np.where(mx > 0, (mx - mn) / np.maximum(mx, 1), 0)
    sat_sel = sub[sat >= np.percentile(sat, 96)]
    satc = sat_sel.mean(axis=0) if len(sat_sel) else med
    f = lambda v: [int(round(t)) for t in v]
    return {"median": f(med), "dark4": f(dark), "sat": f(satc)}


def parse_box(s: str):
    name, coords = s.split(":", 1)
    x1, y1, x2, y2 = (int(v) for v in coords.split(","))
    return name, (x1, y1, x2, y2)


def main() -> int:
    ap = argparse.ArgumentParser(description="poster-vba measure colors")
    ap.add_argument("img")
    ap.add_argument("--top", type=int, default=12)
    ap.add_argument("--box", action="append", default=[])
    args = ap.parse_args()
    img = Image.open(args.img).convert("RGB")
    a = np.asarray(img).astype(np.float32)

    ring = np.zeros(a.shape[:2], bool)
    t = 2
    ring[:t, :] = ring[-t:, :] = True
    ring[:, :t] = ring[:, -t:] = True
    bg = np.median(a[ring].reshape(-1, 3), axis=0)

    res = {
        "image": "%dx%d" % (img.width, img.height),
        "bg_guess": [int(v) for v in bg.round()],
        "dominant": dominant(img, args.top),
        "boxes": {},
    }
    for spec in args.box:
        name, box = parse_box(spec)
        res["boxes"][name] = box_colors(a, box)
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
