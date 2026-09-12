# -*- coding: utf-8 -*-
"""poster-vba 图标素材工具：从原图裁剪 / 抠底去背 / 清晰化放大，输出透明 PNG。

用法:
    python icon_tool.py probe  <img>                     # 看尺寸/边框色/是否带 alpha
    python icon_tool.py check  <img>                     # 判断是否已是透明背景
    python icon_tool.py cut <src> <out.png> [--bbox x1,y1,x2,y2]
            [--keyout auto|none|R,G,B] [--tol 40] [--trim]
            [--scale 2] [--sharpen 0.8]

- --bbox 省略时取整图。
- --keyout auto: 以边界像素的中位色为参考色，从四边洪泛填充删除背景
  （只删与边界连通的背景，图标内部同色区域保留）。适合纯色/浅色底展板。
- --keyout R,G,B: 全图删除接近该颜色的像素（含内部孔洞），适合纯色底图标。
- --tol 颜色容差（欧氏距离，默认 40；背景有轻微渐变时适当加大）。
- --trim 按 alpha>8 的内容包围盒裁掉透明边（留 2px 余量）。
- --scale Lanczos 放大倍数（1~4，模糊小图建议 2~3）。
- --sharpen 锐化强度 0~1（默认 0.8；只锐化 RGB，不破坏 alpha）。

输出: 带 alpha 的 PNG（python-pptx / PowerPoint 均支持透明）。
依赖: Pillow、numpy（隔离 venv 已预装）。
"""
import argparse
import os
import sys

import numpy as np
from PIL import Image, ImageFilter


def _border_ring(a: np.ndarray, t: int = 2) -> np.ndarray:
    mask = np.zeros(a.shape[:2], bool)
    mask[:t, :] = mask[-t:, :] = True
    mask[:, :t] = mask[:, -t:] = True
    return mask


def _ref_color(a: np.ndarray) -> np.ndarray:
    ring = _border_ring(a)
    px = a[ring].reshape(-1, 3).astype(np.float32)
    return np.median(px, axis=0)


def _flood_bg(rgb: np.ndarray, ref: np.ndarray, tol: float) -> np.ndarray:
    """Flood from the borders over pixels within tol of ref. Vectorized."""
    h, w, _ = rgb.shape
    dist = np.sqrt(((rgb.astype(np.float32) - ref) ** 2).sum(axis=2))
    match = dist <= tol
    bg = _border_ring(rgb) & match
    total_match = int(match.sum())
    it = 0
    while True:
        grown = bg.copy()
        grown[1:, :] |= bg[:-1, :]
        grown[:-1, :] |= bg[1:, :]
        grown[:, 1:] |= bg[:, :-1]
        grown[:, :-1] |= bg[:, 1:]
        grown &= match
        it += 1
        if grown.sum() == bg.sum() or it > 6000:
            break
        bg = grown
    # 边框色与内容色太接近导致洪泛没启动 -> 退化为全图色匹配
    if total_match > 0 and bg.sum() < 0.005 * total_match:
        return match
    return bg


def _keyout_solid(rgb: np.ndarray, color, tol: float) -> np.ndarray:
    ref = np.array(color, np.float32)
    dist = np.sqrt(((rgb.astype(np.float32) - ref) ** 2).sum(axis=2))
    return dist <= tol


def cmd_probe(path: str) -> int:
    img = Image.open(path)
    print("size: %dx%d  mode: %s" % (img.width, img.height, img.mode))
    a = np.asarray(img.convert("RGBA")).astype(np.float32)
    rgb, alpha = a[..., :3], a[..., 3]
    ring = _border_ring(a)
    print("border median RGB: %s" % np.median(rgb[ring], axis=0).round(0).astype(int))
    print("alpha: min %d  %%transparent %.1f%%" % (int(alpha.min()), 100.0 * (alpha < 16).mean()))
    return 0


def cmd_check(path: str) -> int:
    img = Image.open(path)
    a = np.asarray(img.convert("RGBA"))
    alpha = a[..., 3]
    transparent = (alpha < 16).mean()
    semi = ((alpha >= 16) & (alpha < 240)).mean()
    print("alpha<16 占 %.1f%%，半透明 %.1f%%" % (100 * transparent, 100 * semi))
    if not img.mode.endswith("A"):
        print("结论: 不带 alpha 通道 -> 需要 keyout 抠底")
        return 1
    if transparent > 0.05:
        print("结论: 已是透明背景素材，可直接用（必要时 --trim 裁边）")
        return 0
    print("结论: 带 alpha 通道但几乎全不透明 -> 仍需 keyout")
    return 1


def cmd_cut(src: str, out: str, bbox, keyout: str, tol: float, trim: bool,
            scale: float, sharpen: float) -> int:
    img = Image.open(src).convert("RGBA")
    if bbox:
        x1, y1, x2, y2 = bbox
        img = img.crop((x1, y1, x2, y2))
    a = np.asarray(img).astype(np.float32)
    rgb = a[..., :3]
    alpha = a[..., 3].copy()

    if keyout == "auto":
        ref = _ref_color(a[..., :3])
        bgmask = _flood_bg(rgb, ref, tol)
        print("keyout auto: 参考色 RGB%s, 抠掉 %.1f%% 像素"
              % (ref.round(0).astype(int), 100.0 * bgmask.mean()))
        if bgmask.mean() > 0.92:
            print("警告: 被抠掉的比例过高，bbox 可能没框住内容或 tol 过大", file=sys.stderr)
        alpha[bgmask] = 0
    elif keyout not in ("none", ""):
        color = tuple(int(v) for v in keyout.split(","))
        bgmask = _keyout_solid(rgb, color, tol)
        print("keyout %s: 抠掉 %.1f%% 像素" % (keyout, 100.0 * bgmask.mean()))
        alpha[bgmask] = 0

    out_img = Image.fromarray(np.dstack([rgb.astype(np.uint8),
                                         alpha.astype(np.uint8)]), "RGBA")
    # 去白边: alpha 先收缩 1px 再轻微羽化
    out_img.putalpha(out_img.getchannel("A")
                     .filter(ImageFilter.MinFilter(3))
                     .filter(ImageFilter.GaussianBlur(0.8)))

    if trim:
        bbox_a = out_img.getchannel("A").point(lambda v: 255 if v > 8 else 0)
        box = bbox_a.getbbox()
        if box:
            pad = 2
            box = (max(0, box[0] - pad), max(0, box[1] - pad),
                   min(out_img.width, box[2] + pad), min(out_img.height, box[3] + pad))
            out_img = out_img.crop(box)

    if scale and scale > 1:
        out_img = out_img.resize((int(out_img.width * scale),
                                  int(out_img.height * scale)), Image.LANCZOS)

    if sharpen and sharpen > 0:
        rgb_only = out_img.convert("RGB")
        rgb_only = rgb_only.filter(ImageFilter.UnsharpMask(
            radius=2, percent=int(140 * sharpen), threshold=2))
        out_img = Image.merge("RGBA", (*rgb_only.split(), out_img.getchannel("A")))

    out_img.save(out, "PNG", optimize=True)
    print("OK %s  %dx%d  (%.1f KB)" % (out, out_img.width, out_img.height,
                                       os.path.getsize(out) / 1024.0))
    return 0


def _parse_bbox(s: str):
    x1, y1, x2, y2 = (int(v) for v in s.split(","))
    return (x1, y1, x2, y2)


def main() -> int:
    import os
    ap = argparse.ArgumentParser(description="poster-vba icon asset tool")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("probe")
    p1.add_argument("img")
    p2 = sub.add_parser("check")
    p2.add_argument("img")
    p3 = sub.add_parser("cut")
    p3.add_argument("src")
    p3.add_argument("out")
    p3.add_argument("--bbox", default=None, help="x1,y1,x2,y2")
    p3.add_argument("--keyout", default="auto", help="auto | none | R,G,B")
    p3.add_argument("--tol", type=float, default=40.0)
    p3.add_argument("--trim", action="store_true")
    p3.add_argument("--scale", type=float, default=1.0)
    p3.add_argument("--sharpen", type=float, default=0.8)
    args = ap.parse_args()
    if args.cmd == "probe":
        return cmd_probe(args.img)
    if args.cmd == "check":
        return cmd_check(args.img)
    if args.cmd == "cut":
        bbox = _parse_bbox(args.bbox) if args.bbox else None
        return cmd_cut(args.src, args.out, bbox, args.keyout, args.tol,
                       args.trim, args.scale, args.sharpen)
    return 2


if __name__ == "__main__":
    sys.exit(main())
