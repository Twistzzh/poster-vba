# -*- coding: utf-8 -*-
"""poster-vba: replay modPoster_Content*.bas (AddBG/AddPicture/AddNode/AddPath/
AddBars/SetPara/SetPartColor) and render a preview PNG side by side with ANY
source figure.

Usage: python render_preview.py <output_dir> [source_image]

This mirrors sci-flowchart-vba/render_preview.py but speaks the poster API:
it imports the exact parser/canvas math from build_poster.py so the rendered
layout matches what build_poster.py (COM or replay) will produce. Output:
  - preview_render.png   the replay alone (canvas-px resolution)
  - preview_compare.png  source (top) vs replay (bottom), same aspect
Dependencies: Pillow (+ build_poster.py in the same scripts/ dir).
"""
import math
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from build_poster import (  # noqa
    load_bas, canvas_geometry, resolve_color, resolve_asset, split_args, num,
)
from PIL import Image, ImageDraw, ImageFont  # noqa

FILL_NONE = ("-1", "")

CJK_FONT_FILES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttf",
]
LATIN_FONT_FILES = [
    r"C:\Windows\Fonts\arial.ttf",
    r"C:\Windows\Fonts\times.ttf",
]
FONT_CACHE = {}


def has_cjk(s):
    return any('\u3000' <= c <= '\u303f' or '\u3400' <= c <= '\u9fff'
               or '\uff00' <= c <= '\uffef' for c in s)


def cjk_font_for(family, bold):
    f = (family or "").lower()
    if "simhei" in f or "黑" in f or "hei" in f:
        return r"C:\Windows\Fonts\simhei.ttf"
    if "simsun" in f or "宋" in f or "sun" in f or "serif" in f or "times" in f:
        return r"C:\Windows\Fonts\simsun.ttc"
    return r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc"


def getfont(pt_px, bold, cjk, family=""):
    key = (round(pt_px, 2), bold, cjk, family.lower())
    if key in FONT_CACHE:
        return FONT_CACHE[key]
    if cjk:
        cands = [cjk_font_for(family, bold)] + list(CJK_FONT_FILES)
    else:
        f = (r"C:\Windows\Fonts\timesbd.ttf" if bold else
             r"C:\Windows\Fonts\times.ttf")
        cands = [f] + list(LATIN_FONT_FILES)
    got = None
    for f in cands:
        try:
            got = ImageFont.truetype(f, max(6, int(round(pt_px))))
            break
        except OSError:
            continue
    FONT_CACHE[key] = got if got is not None else ImageFont.load_default()
    return FONT_CACHE[key]


# ---------------------------------------------------------------- shapes

def _poly(d, pts, fill, linec, lw, closed=True):
    if fill is not None:
        d.polygon(pts, fill=fill)
    if linec is not None:
        seq = pts + [pts[0]] if closed else pts
        w = max(1, int(round(lw)))
        for i in range(len(seq) - 1):
            d.line([seq[i], seq[i + 1]], fill=linec, width=w)


def _dashed_rect(d, x, y, w, h, linec, lw):
    wdt = max(1, int(round(lw)))
    for a, b in (((x, y), (x + w, y)), ((x + w, y), (x + w, y + h)),
                 ((x + w, y + h), (x, y + h)), ((x, y + h), (x, y))):
        (x0, y0), (x1, y1) = a, b
        L = math.hypot(x1 - x0, y1 - y0)
        ux, uy = (x1 - x0) / L, (y1 - y0) / L
        t = 0.0
        while t < L:
            e = min(t + 9, L)
            d.line([(x0 + ux * t, y0 + uy * t), (x0 + ux * e, y0 + uy * e)],
                   fill=linec, width=wdt)
            t = e + 6


def star_points(cx, cy, r_out, r_in, n, rot=-math.pi / 2):
    pts = []
    for i in range(2 * n):
        r = r_out if i % 2 == 0 else r_in
        a = rot + i * math.pi / n
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def heart_points(cx, cy, w, h, n=40):
    pts = []
    for i in range(n + 1):
        t = math.pi - math.pi * 2 * i / n
        x = 16 * math.sin(t) ** 3
        y = (13 * math.cos(t) - 5 * math.cos(2 * t) - 2 * math.cos(3 * t)
             - math.cos(4 * t))
        pts.append((cx + x * w / 32.0, cy - y * h / 32.0))
    return pts


def cloud_points(x, y, w, h):
    pts = []
    cx = x + w / 2
    base = y + h * 0.72
    # bumpy top made of arcs; approximate with overlapping circles union outline
    circles = [
        (x + w * 0.22, base, w * 0.20),
        (x + w * 0.40, y + h * 0.42, w * 0.26),
        (x + w * 0.62, y + h * 0.40, w * 0.24),
        (x + w * 0.80, base, w * 0.19),
        (x + w * 0.50, base + h * 0.10, w * 0.34),
    ]
    # draw filled circles (union) instead of an exact polygon for preview
    return ("cloud", circles, (x, y, w, h))


def draw_shape(d, n, col):
    """Draw a node's fill/line/outline. Returns nothing; text drawn separately."""
    k = n["kind"]
    x, y, w, h = n["x"], n["y"], n["w"], n["h"]
    fill = col(n["fill"])
    linec = col(n["line"])
    lw = n["lw"]
    dash = (n.get("dash") or "LINE_SOLID") != "LINE_SOLID"
    box = [x, y, x + w, y + h]

    if k in ("rect", "rectangle", "box"):
        if fill is not None:
            d.rectangle(box, fill=fill)
        if linec is not None:
            if dash:
                _dashed_rect(d, x, y, w, h, linec, lw)
            else:
                d.rectangle(box, outline=linec, width=max(1, int(round(lw))))
    elif k in ("round_rect", "rounded", "rounded_rect", "stadium", "pill",
               "terminator"):
        r = h / 2 if k in ("stadium", "pill", "terminator") else \
            max(2, min(w, h) * (n["corner"] if n["corner"] > 0 else 0.14))
        if fill is not None:
            d.rounded_rectangle(box, radius=r, fill=fill)
        if linec is not None:
            d.rounded_rectangle(box, radius=r, outline=linec, width=max(1, int(round(lw))))
    elif k in ("oval", "ellipse", "circle"):
        if fill is not None:
            d.ellipse(box, fill=fill)
        if linec is not None:
            d.ellipse(box, outline=linec, width=max(1, int(round(lw))))
    elif k == "diamond":
        _poly(d, [(x + w / 2, y), (x + w, y + h / 2), (x + w / 2, y + h), (x, y + h / 2)],
              fill, linec, lw)
    elif k == "parallelogram":
        off = w * 0.2
        _poly(d, [(x + off, y), (x + w, y), (x + w - off, y + h), (x, y + h)], fill, linec, lw)
    elif k == "trapezoid":
        off = w * 0.2
        _poly(d, [(x, y), (x + w, y), (x + w - off, y + h), (x + off, y + h)], fill, linec, lw)
    elif k == "pentagon":  # double-pointed banner (home-plate-ish)
        tip = min(w, h) * (n["corner"] if n["corner"] > 0 else 0.25)
        _poly(d, [(x + tip, y), (x + w - tip, y), (x + w, y + h / 2),
                  (x + w - tip, y + h), (x + tip, y + h), (x, y + h / 2)], fill, linec, lw)
    elif k == "hexagon":
        tip = min(w, h) * (n["corner"] if n["corner"] > 0 else 0.25)
        _poly(d, [(x + tip, y), (x + w - tip, y), (x + w, y + h / 2),
                  (x + w - tip, y + h), (x + tip, y + h), (x, y + h / 2)], fill, linec, lw)
    elif k == "chevron":
        tip = h * 0.5
        _poly(d, [(x, y), (x + w - tip, y), (x + w, y + h / 2), (x + w - tip, y + h),
                  (x, y + h), (x + tip, y + h / 2)], fill, linec, lw)
    elif k in ("right_arrow", "arrow_right"):
        hd = w * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bh = h * 0.5
        cy = y + h / 2
        _poly(d, [(x, cy - bh / 2), (x + w - hd, cy - bh / 2), (x + w - hd, y),
                  (x + w, cy), (x + w - hd, y + h), (x + w - hd, cy + bh / 2),
                  (x, cy + bh / 2)], fill, linec, lw)
    elif k in ("left_arrow", "arrow_left"):
        hd = w * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bh = h * 0.5
        cy = y + h / 2
        _poly(d, [(x + w, cy - bh / 2), (x + hd, cy - bh / 2), (x + hd, y),
                  (x, cy), (x + hd, y + h), (x + hd, cy + bh / 2),
                  (x + w, cy + bh / 2)], fill, linec, lw)
    elif k in ("up_arrow", "arrow_up"):
        hd = h * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bw = w * 0.5
        cx = x + w / 2
        _poly(d, [(cx - bw / 2, y + h), (cx - bw / 2, y + hd), (x, y + hd),
                  (cx, y), (x + w, y + hd), (cx + bw / 2, y + hd),
                  (cx + bw / 2, y + h)], fill, linec, lw)
    elif k in ("down_arrow", "arrow_down"):
        hd = h * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bw = w * 0.5
        cx = x + w / 2
        _poly(d, [(cx - bw / 2, y), (cx - bw / 2, y + h - hd), (x, y + h - hd),
                  (cx, y + h), (x + w, y + h - hd), (cx + bw / 2, y + h - hd),
                  (cx + bw / 2, y)], fill, linec, lw)
    elif k in ("can", "cylinder"):
        if fill is not None:
            d.rectangle([x, y + h * 0.12, x + w, y + h * 0.88], fill=fill)
        if fill is not None:
            d.ellipse([x, y, x + w, y + h * 0.24], fill=fill)
        if fill is not None:
            d.ellipse([x, y + h * 0.76, x + w, y + h], fill=fill)
        if linec is not None:
            wdt = max(1, int(round(lw)))
            d.ellipse([x, y, x + w, y + h * 0.24], outline=linec, width=wdt)
            d.line([(x, y + h * 0.12), (x, y + h * 0.88)], fill=linec, width=wdt)
            d.line([(x + w, y + h * 0.12), (x + w, y + h * 0.88)], fill=linec, width=wdt)
            d.ellipse([x, y + h * 0.76, x + w, y + h], outline=linec, width=wdt)
    elif k == "document":
        fg = [(x, y), (x + w * 0.7, y), (x + w, y + h * 0.3), (x + w, y + h),
              (x, y + h)]
        _poly(d, fg, fill, None, lw)
        if linec is not None:
            d.line([(x + w * 0.7, y), (x + w * 0.7, y + h * 0.3), (x + w, y + h * 0.3)],
                   fill=linec, width=max(1, int(round(lw))))
    elif k == "arc":
        if linec is not None:
            d.arc(box, 180, 360, fill=linec, width=max(1, int(round(lw))))
        if fill is not None:
            d.pieslice(box, 180, 360, fill=fill)
    elif k in ("star4", "star5", "star6", "star8"):
        npts = int(k[-1])
        rin = min(w, h) * (n["corner"] if 0 < n["corner"] < 1 else 0.45)
        rout = min(w, h) / 2.0
        cx, cy = x + w / 2, y + h / 2
        _poly(d, star_points(cx, cy, rout, rin, npts), fill, linec, lw)
    elif k in ("heart",):
        cx = x + w / 2
        cy = y + h / 2
        _poly(d, heart_points(cx, cy, w, h), fill, linec, lw)
    elif k in ("cloud",):
        _, circles, _ = cloud_points(x, y, w, h)
        if fill is not None:
            for (cx, cy, r) in circles:
                d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill)
        if linec is not None:
            wdt = max(1, int(round(lw)))
            for (cx, cy, r) in circles:
                d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=linec, width=wdt)
    elif k in ("sun",):
        cx, cy = x + w / 2, y + h / 2
        r = min(w, h) * 0.30
        if fill is not None:
            d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill)
        if linec is not None:
            wdt = max(1, int(round(lw)))
            d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=linec, width=wdt)
            for i in range(8):
                a = i * math.pi / 4
                d.line([(cx + (r + 2) * math.cos(a), cy + (r + 2) * math.sin(a)),
                        (cx + (r + w * 0.20) * math.cos(a), cy + (r + w * 0.20) * math.sin(a))],
                       fill=linec, width=wdt)
    elif k in ("moon",):
        cx, cy = x + w / 2, y + h / 2
        r = min(w, h) / 2.0
        if fill is not None:
            d.ellipse([x, y, x + w, y + h], fill=fill)
            d.ellipse([cx + r * 0.5, cy - r, cx + r * 1.5, cy + r],
                      fill=(255, 255, 255))
    elif k in ("lightning", "bolt"):
        pts = [(x + w * 0.55, y), (x + w * 0.25, y + h * 0.55),
               (x + w * 0.45, y + h * 0.55), (x + w * 0.30, y + h),
               (x + w * 0.80, y + h * 0.40), (x + w * 0.55, y + h * 0.40)]
        _poly(d, pts, fill, linec, lw)
    elif k in ("smiley",):
        cx, cy = x + w / 2, y + h / 2
        r = min(w, h) / 2.0
        if fill is not None:
            d.ellipse([x, y, x + w, y + h], fill=fill)
        if linec is not None:
            wdt = max(1, int(round(lw)))
            d.ellipse([x, y, x + w, y + h], outline=linec, width=wdt)
            d.ellipse([cx - r * 0.4, cy - r * 0.3, cx - r * 0.1, cy - r * 0.0],
                      outline=linec, width=wdt)
            d.ellipse([cx + r * 0.1, cy - r * 0.3, cx + r * 0.4, cy - r * 0.0],
                      outline=linec, width=wdt)
            d.arc([cx - r * 0.5, cy, cx + r * 0.5, cy + r * 0.6], 20, 160,
                  fill=linec, width=wdt)
    elif k in ("donut",):
        cx, cy = x + w / 2, y + h / 2
        r = min(w, h) / 2.0
        if fill is not None:
            d.ellipse([x, y, x + w, y + h], fill=fill)
            d.ellipse([cx - r * 0.5, cy - r * 0.5, cx + r * 0.5, cy + r * 0.5],
                      fill=(255, 255, 255))
    elif k in ("wave", "double_wave"):
        amp = h * 0.22
        mid = y + h / 2
        pts = []
        nseg = 60
        for i in range(nseg + 1):
            px = x + w * i / nseg
            py = mid + amp * math.sin(2 * math.pi * (i / nseg) * (2 if k == "double_wave" else 1))
            pts.append((px, py))
        if fill is not None:
            pts2 = list(pts) + [(x + w, y + h), (x, y + h)]
            d.polygon(pts2, fill=fill)
        if linec is not None:
            d.line(pts, fill=linec, width=max(1, int(round(lw))))
    elif k in ("callout", "round_callout", "bubble", "oval_callout"):
        r = min(w, h) * 0.18
        if fill is not None:
            d.rounded_rectangle(box, radius=r, fill=fill)
        if linec is not None:
            d.rounded_rectangle(box, radius=r, outline=linec, width=max(1, int(round(lw))))
        # tail
        tcx = x + w * 0.5
        tcy = y + h
        if fill is not None:
            d.polygon([(tcx - w * 0.10, tcy), (tcx + w * 0.10, tcy),
                       (tcx + w * 0.02, tcy + h * 0.18)], fill=fill)
    else:  # unknown -> rounded rect fallback
        r = min(w, h) * 0.14
        if fill is not None:
            d.rounded_rectangle(box, radius=r, fill=fill)
        if linec is not None:
            d.rounded_rectangle(box, radius=r, outline=linec, width=max(1, int(round(lw))))


def draw_text(d, n, col, px_per_pt):
    if not n["text"]:
        return
    cjk = has_cjk(n["text"])
    fam = n.get("fname") or n.get("_font") or ""
    f = getfont(n["pt"] * px_per_pt, n["bold"], cjk, fam)
    tc = col(n["fc"]) or (31, 31, 31)
    left = n.get("align") == "left"
    ml_px = (n.get("ml") or 0)
    lines = []
    for para in n["text"].split("\n"):
        units = list(para) if cjk else para.split()
        sep = "" if cjk else " "
        lim = n["w"] - 6 - (ml_px if left else 0)
        cur = ""
        for wd in units:
            t = (cur + sep + wd).strip() if sep else (cur + wd)
            if (d.textlength(t, font=f) <= lim) or not cur:
                cur = t
            else:
                lines.append(cur)
                cur = wd
        lines.append(cur)
    lh = n["pt"] * px_per_pt
    total = lh * len(lines)
    ty = n["y"] + n["h"] / 2 - total / 2
    prefix = n.get("prefix")
    for li, ln in enumerate(lines):
        if left:
            x0 = n["x"] + 3 + ml_px
            d.text((x0, ty), ln, font=f, fill=tc, anchor="la")
        elif prefix and prefix[0] > 0 and li == 0 and prefix[0] < len(ln):
            pre, rest = ln[:prefix[0]], ln[prefix[0]:]
            x0 = n["x"] + n["w"] / 2 - d.textlength(ln, font=f) / 2
            fb = getfont(n["pt"] * px_per_pt, True, cjk, fam)
            pcol = col(prefix[1]) or tc
            d.text((x0, ty), pre, font=fb, fill=pcol, anchor="la")
            d.text((x0 + d.textlength(pre, font=fb), ty), rest, font=f, fill=tc, anchor="la")
        else:
            d.text((n["x"] + n["w"] / 2, ty), ln, font=f, fill=tc, anchor="ma")
        ty += lh


def _render_tex_png(tex: str, size_pt: float, col_rgb, px_per_pt: float):
    """Render a LaTeX string via matplotlib mathtext; returns cropped RGBA image
    scaled for the canvas (1pt == px_per_pt px). Raises on any failure."""
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    fig = figure.Figure(figsize=(8, 2.5), dpi=72)
    FigureCanvasAgg(fig)
    fig.patch.set_alpha(0.0)
    s = tex.replace("\n", " ").strip()
    s = s.replace("\\,", " ").replace("\\;", " ").replace("\\!", "")
    s = re.sub(r"\s+", " ", s)
    col01 = tuple(c / 255.0 for c in (col_rgb or (31, 31, 31)))
    fig.text(0.5, 0.5, "$%s$" % s, fontsize=size_pt, color=col01,
             math_fontfamily="stix", ha="center", va="center")
    fig.canvas.draw()
    W_, H_ = fig.canvas.get_width_height()
    buf = bytes(fig.canvas.buffer_rgba())
    arr = Image.frombytes("RGBA", (W_, H_), buf)
    alpha = arr.split()[3]
    bbox = alpha.getbbox()
    if not bbox:
        return None
    arr = arr.crop(bbox)
    k = px_per_pt
    nw, nh = max(1, int(arr.width * k)), max(1, int(arr.height * k))
    return arr.resize((nw, nh), Image.LANCZOS)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    out_dir = os.path.abspath(sys.argv[1])
    src = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else None

    if not os.path.isdir(out_dir):
        print("FAIL: 目录不存在: %s" % out_dir)
        return 2

    pal, scal, ops, decor = load_bas(out_dir)
    g = canvas_geometry(scal)
    CW, CH = int(g["cw"]), int(g["ch"])
    px_per_pt = 1.0 / (g["s"] * 72.0) if g["s"] > 0 else 1.0

    img = Image.new("RGB", (CW, CH), (255, 255, 255))
    d = ImageDraw.Draw(img)

    def col(c):
        c = (c or "").strip()
        if c in FILL_NONE:
            return None
        rgb = resolve_color(c, pal)
        return rgb

    # build node lookups for decor (SetPara / SetPartColor)
    nodes = {}
    for o in ops:
        if o["op"] == "node":
            nodes[o["id"]] = o
            o["_font"] = g["font"]

    for line in decor:
        if line.startswith("SetPara "):
            a = split_args(line[7:].strip())
            if len(a) < 3:
                continue
            nid = a[1].strip('"')
            n = nodes.get(nid)
            if n:
                n["align"] = a[2].strip('"').lower()
                n["ml"] = num(a[3]) if len(a) > 3 and a[3] else 0.0
        elif line.startswith("SetPartColor "):
            a = split_args(line[12:].strip())
            if len(a) < 3:
                continue
            nid = a[1].strip('"')
            n = nodes.get(nid)
            if n:
                nc = int(num(a[2])) if len(a) > 2 and a[2] else 0
                ctoken = a[3] if len(a) > 3 else "RGB(200,30,40)"
                n["prefix"] = (nc, ctoken)

    # draw in call order (z-order parity with VBA / replay)
    n_pic = 0
    for o in ops:
        if o["op"] == "bg":
            c1 = col(o["c1"]) or (255, 255, 255)
            c2 = col(o["c2"])
            if c2 is None:
                d.rectangle([0, 0, CW, CH], fill=c1)
            else:
                for yy in range(CH):
                    t = yy / max(1, CH - 1)
                    r = int(c1[0] + (c2[0] - c1[0]) * t)
                    gg = int(c1[1] + (c2[1] - c1[1]) * t)
                    bb = int(c1[2] + (c2[2] - c1[2]) * t)
                    d.line([(0, yy), (CW, yy)], fill=(r, gg, bb))
        elif o["op"] == "picture":
            resolved = resolve_asset(o["file"], out_dir, g["assets"])
            if resolved is None:
                pw = o["w"] if o["w"] > 0 else 120.0
                ph = o["h"] if o["h"] > 0 else 90.0
                d.rectangle([o["x"], o["y"], o["x"] + pw, o["y"] + ph],
                            fill=(238, 238, 238), outline=(170, 170, 170))
                d.text((o["x"] + 4, o["y"] + 4), "IMG %s" % o["file"],
                       fill=(120, 120, 120), anchor="la")
                print("  preview: missing asset %r -> placeholder" % o["file"])
            else:
                try:
                    im = Image.open(resolved).convert("RGBA")
                    if o["w"] > 0 and o["h"] > 0:
                        im = im.resize((int(o["w"]), int(o["h"])))
                    img.paste(im, (int(o["x"]), int(o["y"])), im)
                    n_pic += 1
                except Exception as exc:
                    print("  preview: add_picture failed %r -> %s" % (o["file"], exc))
        elif o["op"] == "node":
            draw_shape(d, o, col)
            draw_text(d, o, col, px_per_pt)
        elif o["op"] == "path":
            c = col(o["color"]) or (0, 0, 0)
            wdt = max(1, int(round(o["lw"])))
            is_dash = (o.get("dash") or "LINE_SOLID") != "LINE_SOLID"
            for i in range(len(o["pts"]) - 1):
                a, b = o["pts"][i], o["pts"][i + 1]
                if is_dash:
                    L = math.hypot(b[0] - a[0], b[1] - a[1])
                    ux, uy = (b[0] - a[0]) / L, (b[1] - a[1]) / L
                    t = 0.0
                    while t < L:
                        e = min(t + 9, L)
                        d.line([(a[0] + ux * t, a[1] + uy * t),
                                (a[0] + ux * e, a[1] + uy * e)], fill=c, width=wdt)
                        t = e + 6
                else:
                    d.line([a, b], fill=c, width=wdt)
            if o["arrow"] and len(o["pts"]) >= 2:
                a, b = o["pts"][-2], o["pts"][-1]
                ang = math.atan2(b[1] - a[1], b[0] - a[0])
                d.polygon([b,
                           (b[0] + 13 * math.cos(ang + 2.7), b[1] + 13 * math.sin(ang + 2.7)),
                           (b[0] + 13 * math.cos(ang - 2.7), b[1] + 13 * math.sin(ang - 2.7))],
                          fill=c)
        elif o["op"] == "formula":
            tc = col(o["ctok"]) or (31, 31, 31)
            rendered = None
            try:
                rendered = _render_tex_png(o["tex"], o["pt"], tc, px_per_pt)
            except Exception:
                rendered = None
            if rendered is not None:
                px = int(o["x"] + o["w"] / 2 - rendered.width / 2)
                py = int(o["y"] + o["h"] / 2 - rendered.height / 2)
                img.paste(rendered, (px, py), rendered)
            else:
                cjk = has_cjk(o["fb"])
                fnt = getfont(max(6, o["pt"] * 0.6 * px_per_pt), False, cjk, "")
                d.text((o["x"] + o["w"] / 2, o["y"] + o["h"] / 2), o["fb"],
                       font=fnt, fill=tc, anchor="mm")

    img.save(os.path.join(out_dir, "preview_render.png"))
    print("preview_render: %dx%d  pictures=%d" % (CW, CH, n_pic))

    if src and os.path.exists(src):
        s = Image.open(src).convert("RGB").resize((CW, CH))
        cmp_img = Image.new("RGB", (CW, CH * 2 + 16), (225, 225, 225))
        cmp_img.paste(s, (0, 0))
        cmp_img.paste(img, (0, CH + 16))
        sc = 1500.0 / CW
        cmp_img = cmp_img.resize((1500, int((CH * 2 + 16) * sc)), Image.LANCZOS)
        cmp_img.save(os.path.join(out_dir, "preview_compare.png"))
        print("preview_compare: 上=源图 下=渲染 (1500px 宽)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
