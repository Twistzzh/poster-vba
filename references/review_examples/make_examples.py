# -*- coding: utf-8 -*-
"""生成 poster-vba 评审用的正负样本图（origin.png 干净正样本 / failed.png 典型失败）。

仅用于复现 review_examples 里的示例图，与任何真实用户任务无关。
用法: python make_examples.py
"""
import math
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
W, H = 1200, 675

FONT_DIR = r"C:\Windows\Fonts"
_C = {}


def font(pt, bold=False, cjk=True):
    key = (pt, bold, cjk)
    if key in _C:
        return _C[key]
    cand = ([os.path.join(FONT_DIR, "msyhbd.ttc")] if bold else
            [os.path.join(FONT_DIR, "msyh.ttc")]) if cjk else \
           ([os.path.join(FONT_DIR, "arialbd.ttf")] if bold else
            [os.path.join(FONT_DIR, "arial.ttf")])
    got = None
    for f in cand:
        try:
            got = ImageFont.truetype(f, int(pt))
            break
        except OSError:
            continue
    _C[key] = got or ImageFont.load_default()
    return _C[key]


def rr(d, box, r, fill=None, outline=None, width=1):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=width)


def star(d, cx, cy, r, n, fill, outline=None):
    pts = []
    for i in range(2 * n):
        rad = r if i % 2 == 0 else r * 0.45
        a = -math.pi / 2 + i * math.pi / n
        pts.append((cx + rad * math.cos(a), cy + rad * math.sin(a)))
    d.polygon(pts, fill=fill, outline=outline)


def heart(d, cx, cy, w, h, fill):
    pts = []
    for i in range(60):
        t = math.pi - math.pi * 2 * i / 60
        x = 16 * math.sin(t) ** 3
        y = (13 * math.cos(t) - 5 * math.cos(2 * t) - 2 * math.cos(3 * t)
             - math.cos(4 * t))
        pts.append((cx + x * w / 32.0, cy - y * h / 32.0))
    d.polygon(pts, fill=fill)


def battery(d, x, y, w, h, fill, text=""):
    rr(d, (x, y, x + w, y + h), 14, fill=fill, outline=(40, 120, 70), width=3)
    d.rectangle((x + w * 0.38, y - 12, x + w * 0.62, y + 4), fill=(40, 120, 70))
    if text:
        f = font(26, True, False)
        d.text((x + w / 2, y + h / 2), text, font=f, fill=(255, 255, 255),
               anchor="mm")


def bars(d, x, y, w, h, cats, vals, bar_c, axis_c, txt_c, title=""):
    if title:
        f = font(18, True)
        d.text((x, y), title, font=f, fill=txt_c, anchor="la")
        y += 30
    vmax = max(vals) * 1.1
    plot_h = h * 0.7
    base = y + plot_h
    slot = w / len(vals)
    for i, v in enumerate(vals):
        bx = x + slot * i + 0.15 * slot
        bh = (v / vmax) * plot_h
        rr(d, (bx, base - bh, bx + 0.7 * slot, base), 4, fill=bar_c)
        f = font(14, False, False)
        d.text((bx + 0.35 * slot, base - bh - 18), "%.1f" % v, font=f,
               fill=txt_c, anchor="mm")
        fc = font(15)
        d.text((x + slot * i + slot / 2, base + h * 0.08), cats[i], font=fc,
               fill=txt_c, anchor="mm")
    d.line([(x, base), (x + w, base)], fill=axis_c, width=2)


def draw_layout(d, faded=False, overflow=False, missing_icon=False,
                star_on_title=False, wrong_bars=False, panel_white=False):
    # 背景渐变
    c_top = (235, 220, 200) if faded else (255, 248, 235)
    c_bot = (225, 205, 175) if faded else (255, 236, 210)
    for yy in range(H):
        t = yy / H
        d.line([(0, yy), (W, yy)],
               fill=(int(c_top[0] + (c_bot[0] - c_top[0]) * t),
                     int(c_top[1] + (c_bot[1] - c_top[1]) * t),
                     int(c_top[2] + (c_bot[2] - c_top[2]) * t)))

    # 标题条
    title_fill = (120, 140, 175) if faded else (64, 132, 220)
    rr(d, (60, 40, 1140, 100), 18, fill=title_fill)
    f = font(30, True)
    d.text((600, 70), "2026 城市环保创意大赛 展板", font=f, fill=(255, 255, 255),
           anchor="mm")

    # 面板 A
    pa_fill = (245, 245, 245) if panel_white else (224, 239, 252)
    rr(d, (60, 130, 420, 360), 14, fill=pa_fill)
    fh = font(20, True)
    d.text((80, 150), "研究背景", font=fh, fill=(31, 78, 150 if not faded else 110),
           anchor="la")
    fb = font(15)
    body = ["城市垃圾分类推行三年来，居民", "参与率提升明显，但投放准确",
            "率仍不足六成。本作品设计", "一套智能督导装置实时纠正。"]
    if overflow:
        body = body + ["（这一行已经超出面板底", "边，属于典型失败）"]
    ty = 200
    for ln in body:
        d.text((80, ty), ln, font=fb, fill=(51, 51, 51), anchor="la")
        ty += 34

    # 面板 B
    pb_fill = (247, 247, 247) if panel_white else (255, 238, 220)
    rr(d, (450, 130, 840, 360), 14, fill=pb_fill)
    d.text((470, 150), "试点数据", font=fh, fill=(200, 100, 20), anchor="la")
    if wrong_bars:
        bars(d, 470, 200, 350, 260, ["调研", "设计", "制作", "答辩"],
             [9.0, 1.2, 7.5, 0.3], (120, 140, 175), (150, 150, 150), (51, 51, 51))
    else:
        bars(d, 470, 200, 350, 260, ["调研", "设计", "制作", "答辩"],
             [3.2, 5.1, 4.4, 2.8], (64, 132, 220), (150, 150, 150), (51, 51, 51))

    # 面板 C
    pc_fill = (245, 245, 245) if panel_white else (232, 245, 233)
    rr(d, (60, 410, 580, 640), 14, fill=pc_fill)
    d.text((80, 430), "成果亮点", font=fh, fill=(56, 128, 82), anchor="la")
    fc = font(15)
    d.text((80, 480), "识别准确率 94.2%，督导响应 0.8 秒。", font=fc,
           fill=(51, 51, 51), anchor="la")
    d.text((80, 520), "已在 2 个社区完成 6 周试点。", font=fc,
           fill=(51, 51, 51), anchor="la")

    # 电池图标
    if missing_icon:
        rr(d, (700, 430, 800, 560), 12, fill=(238, 238, 238),
           outline=(170, 170, 170), width=2)
        d.text((750, 495), "IMG battery.png", font=font(13), fill=(120, 120, 120),
               anchor="mm")
    else:
        battery(d, 700, 430, 100, 130, (76, 175, 110), "94%")

    # 数据图占位框
    rr(d, (840, 430, 980, 560), 10, fill=(255, 255, 255),
       outline=(150, 150, 150), width=2)
    d.text((910, 495), "数据图\n占位", font=font(15), fill=(120, 120, 120),
           anchor="mm")

    # 装饰星星 + 爱心
    if star_on_title:
        star(d, 1080, 70, 26, 5, (150, 150, 150))  # 压在标题条上（遮挡）
    else:
        star(d, 1080, 70, 26, 5, (255, 205, 66))
        star(d, 1140, 130, 18, 5, (255, 205, 66))
    heart(d, 1110, 600, 42, 40, (240, 98, 146))

    # 流程箭头 A->B
    arrow_c = (120, 140, 175) if faded else (64, 132, 220)
    d.line([(430, 300), (450, 300)], fill=arrow_c, width=4)
    d.polygon([(450, 300), (438, 294), (438, 306)], fill=arrow_c)


def main():
    # 正样本
    img = Image.new("RGB", (W, H), (255, 255, 255))
    draw_layout(ImageDraw.Draw(img))
    img.save(os.path.join(HERE, "origin.png"))
    # 负样本：褪色 / 面板过淡 / 文字溢出 / 图标占位 / 星星遮标题 / 柱图错位
    img2 = Image.new("RGB", (W, H), (255, 255, 255))
    draw_layout(ImageDraw.Draw(img2), faded=True, overflow=True,
                missing_icon=True, star_on_title=True, wrong_bars=True,
                panel_white=True)
    img2.save(os.path.join(HERE, "failed.png"))
    print("wrote origin.png + failed.png (%dx%d)" % (W, H))


if __name__ == "__main__":
    main()
