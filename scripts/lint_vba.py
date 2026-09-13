# -*- coding: utf-8 -*-
"""poster-vba: 对生成的 .bas 做结构自检（导入 PowerPoint 前把能查的错都查掉）。

用法: python lint_vba.py <输出目录>

检查项:
  - modPoster_Engine.bas 必须存在且原样附带
  - 内容模块声明全部 11 个几何/字体/素材常量
    (CANVAS_W_PX, CANVAS_H_PX, PX_TO_IN, OFFSET_X_IN, OFFSET_Y_IN,
     SLIDE_W_IN, SLIDE_H_IN, CUSTOM_SIZE, FONT_NAME, FONT_NAME_CN, ASSET_DIR)
  - 存在 Public Sub DrawAll(ByVal sld As Slide) 且调用的 DrawContentN 都有定义
  - AddNode / AddPicture / AddBG / AddBars / AddPath 无空实参（漏写导致 ,,）
  - kind 合法（含卡通族 star/heart/cloud/sun/...）
  - 颜色实参要么引用已定义的调色板常量，要么 RGB(r,g,b) 字面量，要么 -1
  - 每个 .bas 为 UTF-8 / GBK 编码且行尾为 CRLF
退出码: 0 = 通过(允许 WARN), 1 = 有 ERROR。
"""
import io
import os
import re
import sys

GEOM = ["CANVAS_W_PX", "CANVAS_H_PX", "PX_TO_IN", "OFFSET_X_IN", "OFFSET_Y_IN",
        "SLIDE_W_IN", "SLIDE_H_IN", "CUSTOM_SIZE", "FONT_NAME", "FONT_NAME_CN",
        "ASSET_DIR"]
ENGINE_FILE = "modPoster_Engine.bas"
ENGINE_CONST = ["TAG_NAME", "ARROWHEAD_TRIANGLE", "LINE_SOLID", "LINE_DASH",
                "LINE_DOT", "LINE_DASHDOT"]
# 内容模块里出现的引擎过程（跨模块调用合法）
ENGINE_SUBS = ["AddNode", "AddPicture", "AddBG", "AddBars", "AddPath",
               "SetPara", "SetPartColor", "BringToFront", "DrawAll",
               "RemovePoster", "BuildPoster", "ShapeById"]

KINDS = {
    "rect", "rectangle", "box", "round_rect", "rounded", "rounded_rect",
    "oval", "ellipse", "circle", "diamond", "decision", "parallelogram",
    "trapezoid", "pentagon", "pentagon_reg", "hexagon", "stadium", "pill",
    "terminator", "can", "cylinder", "document", "right_arrow", "arrow_right",
    "left_arrow", "arrow_left", "up_arrow", "arrow_up", "down_arrow",
    "arrow_down", "chevron", "arc", "star4", "star5", "star6", "star8",
    "heart", "cloud", "sun", "moon", "lightning", "bolt", "smiley", "donut",
    "wave", "double_wave", "callout", "round_callout", "bubble", "oval_callout",
}


def split_args(s):
    out, cur, depth, q = [], [], 0, False
    for ch in s:
        if ch == '"':
            q = not q
        elif not q and ch == "(":
            depth += 1
        elif not q and ch == ")":
            depth -= 1
        if ch == "," and depth == 0 and not q:
            out.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    cur = "".join(cur).strip()
    if cur:
        out.append(cur)
    return out


def strip_vba_comment(s):
    """Cut a trailing VBA comment: first apostrophe outside string literals."""
    q = False
    for i, ch in enumerate(s):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return s[:i]
    return s


def check_call(a, fn, ln, errs, warns):
    """Validate the arguments of an Add* call (after sld has been dropped)."""
    if any(x == "" for x in a):
        errs.append("%s:%d 存在空实参（可能漏写导致 ,,）" % (fn, ln))


def main(d):
    errs, warns, info = [], [], []
    if not os.path.isdir(d):
        print("FAIL: 目录不存在: %s" % d)
        return 1
    files = sorted(f for f in os.listdir(d) if f.lower().endswith(".bas"))
    if not files:
        errs.append("目录里没有任何 .bas 文件")
        _emit(errs, warns, info)
        return 1

    if ENGINE_FILE not in files:
        errs.append("缺少固定引擎 %s（必须原样附带）" % ENGINE_FILE)

    mods = {}
    for fn in files:
        raw = open(os.path.join(d, fn), "rb").read()
        t = None
        enc = None
        for enc in ("utf-8-sig", "gbk"):
            try:
                t = raw.decode(enc)
                break
            except UnicodeDecodeError:
                pass
        if t is None:
            t = raw.decode("latin-1")
            errs.append("%s: 非 UTF-8 也非 GBK, 导入 VBE 可能乱码" % fn)
        elif enc == "gbk":
            warns.append("%s: GBK 编码中文（中文版 Windows VBE 可正常导入；"
                         "用 build_poster.py 出图则无需此步）" % fn)
        if b"\r\n" not in raw:
            warns.append("%s: 行尾不是 CRLF（回放/导入可能错位）" % fn)
        mods[fn] = t

    # 常量与子过程收集（跨全部模块）
    consts, subs = set(), {}
    for fn, t in mods.items():
        for m in re.finditer(r"Public\s+Const\s+(\w+)\s+As\s+\w+\s*=\s*", t, re.I):
            consts.add(m.group(1))
        for m in re.finditer(r"^\s*(?:Public\s+|Private\s+)?Sub\s+(\w+)\s*\(",
                             t, re.M):
            subs[m.group(1)] = fn

    if "DrawAll" not in subs:
        errs.append("缺少 Public Sub DrawAll(ByVal sld As Slide)")
    missing = [g for g in GEOM if g not in consts]
    if missing:
        errs.append("缺几何/字体/素材常量: %s" % ", ".join(missing))

    # 内容模块里调用了 DrawContentN / DrawBackground 等必须都已定义
    for fn, t in mods.items():
        if fn == ENGINE_FILE:
            continue
        for m in re.finditer(r"^\s{4}([A-Z]\w+)\s+sld$", t, re.M):
            name = m.group(1)
            if name in ("DrawAll",):
                continue
            if name not in subs and name not in ENGINE_SUBS:
                errs.append("%s: 调用了未定义的子过程 %s" % (fn, name))

    # 颜色实参必须可解析：定义过的常量 / RGB(...) / -1
    color_ok = lambda tok: (tok in consts or tok in ENGINE_CONST
                            or re.match(r"RGB\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)",
                                        tok, re.I) is not None
                            or tok in ("-1", ""))

    n_node = n_pic = n_bg = n_bars = n_path = 0
    for fn, t in mods.items():
        if fn == ENGINE_FILE:
            continue  # 引擎内部 AddNode/AddBars 展开是合法的，不查内容模块规则
        logical = re.sub(r"_\s*\r?\n\s*", " ", t)
        for ln, line in enumerate(logical.splitlines(), 1):
            s = strip_vba_comment(line).strip()  # 去掉 VBA 行尾注释后再解析
            if s.startswith("AddNode "):
                n_node += 1
                a = split_args(s[8:])[1:]  # 去掉首参 sld
                if len(a) < 10:
                    errs.append("%s:%d AddNode 实参不足（至少 10，含 text 槽）" % (fn, ln))
                if len(a) > 18:
                    warns.append("%s:%d AddNode 实参过多 (%d)" % (fn, ln, len(a)))
                check_call(a, fn, ln, errs, warns)
                if len(a) > 1:
                    kind = a[1].strip('"')
                    if kind and kind not in KINDS:
                        errs.append("%s:%d 未知形状 kind=%s" % (fn, ln, kind))
                for idx in (6, 7):  # fill, line
                    if len(a) > idx and a[idx] and not color_ok(a[idx]):
                        errs.append("%s:%d 未定义的颜色实参 %s" % (fn, ln, a[idx]))
                if len(a) > 14 and a[14] and not color_ok(a[14]):  # fontC
                    errs.append("%s:%d 未定义的文字颜色 %s" % (fn, ln, a[14]))
            elif s.startswith("AddPicture "):
                n_pic += 1
                a = split_args(s[11:])[1:]
                if len(a) != 6:
                    errs.append("%s:%d AddPicture 实参应为 6，实为 %d" % (fn, ln, len(a)))
                check_call(a, fn, ln, errs, warns)
            elif s.startswith("AddBG "):
                n_bg += 1
                a = split_args(s[6:])[1:]
                if len(a) != 3:
                    errs.append("%s:%d AddBG 实参应为 3，实为 %d" % (fn, ln, len(a)))
                check_call(a, fn, ln, errs, warns)
                for idx in (1, 2):  # c1, c2
                    if len(a) > idx and a[idx] and not color_ok(a[idx]):
                        errs.append("%s:%d 未定义的背景色 %s" % (fn, ln, a[idx]))
            elif s.startswith("AddBars "):
                n_bars += 1
                a = split_args(s[8:])[1:]
                if len(a) < 11:
                    errs.append("%s:%d AddBars 实参不足（至少 11）" % (fn, ln))
                if len(a) > 14:
                    warns.append("%s:%d AddBars 实参过多 (%d)" % (fn, ln, len(a)))
                check_call(a, fn, ln, errs, warns)
                for idx in (8, 9, 10):  # barC, lineC, txtC
                    if len(a) > idx and a[idx] and not color_ok(a[idx]):
                        errs.append("%s:%d 未定义的柱图颜色 %s" % (fn, ln, a[idx]))
            elif s.startswith("AddPath "):
                n_path += 1
                a = split_args(s[8:])[1:]
                if len(a) != 6:
                    errs.append("%s:%d AddPath 实参应为 6，实为 %d" % (fn, ln, len(a)))
                check_call(a, fn, ln, errs, warns)
                pts = a[1].strip('"').split(";") if len(a) > 1 else []
                for p in pts:
                    if p and len(p.split(",")) != 2:
                        errs.append("%s:%d AddPath 坐标段格式错误: %s" % (fn, ln, p))
                for idx in (2,):  # color (lineC)
                    if len(a) > idx and a[idx] and not color_ok(a[idx]):
                        errs.append("%s:%d 未定义的连线颜色 %s" % (fn, ln, a[idx]))
                if len(a) > 4 and a[4] not in ENGINE_CONST and not a[4].isdigit():
                    errs.append("%s:%d 非法 dash=%s" % (fn, ln, a[4]))
                if len(a) > 5 and a[5] not in ("True", "False"):
                    errs.append("%s:%d arrowEnd 非布尔: %s" % (fn, ln, a[5]))
            if len(line) > 900:
                warns.append("%s:%d 行过长 (%d 字符)" % (fn, ln, len(line)))

    info.append("模块: %s" % ", ".join(sorted(mods)))
    info.append("Sub: %d 个 | AddNode %d | AddPicture %d | AddBG %d | "
                "AddBars %d | AddPath %d"
                % (len(subs), n_node, n_pic, n_bg, n_bars, n_path))
    _emit(errs, warns, info)
    return 1 if errs else 0


def _emit(errs, warns, info):
    print("== INFO ==")
    for i in info:
        print("  " + i)
    print("== WARN (%d) ==" % len(warns))
    for w in warns:
        print("  " + w)
    print("== ERROR (%d) ==" % len(errs))
    for e in errs:
        print("  " + e)
    if errs:
        print("RESULT: FAIL (%d err / %d warn) -- 修 .bas 后重跑" % (len(errs), len(warns)))
    else:
        print("RESULT: PASS (%d warn)" % len(warns))


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "."))
