# -*- coding: utf-8 -*-
"""poster-vba step 5 -- turn the delivered .bas modules into a real .pptx.

Usage:
    python build_poster.py <output_dir> [--out name.pptx] [--no-run]
                           [--replay] [--export-png preview.png]

Two paths, tried in this order:

  A. COM path (preferred, Windows + PowerPoint installed)
     Import modPoster_Engine + modPoster_Content* into a blank presentation
     and actually invoke BuildPoster(), so PowerPoint itself draws everything
     (shapes, pictures, gradients). --export-png additionally exports slide 1
     as a PNG for the visual review.

  B. Replay path (fallback, no Office needed)
     Parse AddBG / AddPicture / AddNode / AddPath / AddBars / SetPara /
     SetPartColor calls out of the .bas and emit equivalent native shapes via
     python-pptx. Draw order == call order, so z-order matches the VBA path.

Exits non-zero if neither path produced a slide with shapes.
"""
from __future__ import annotations

import argparse
import os
import re
import sys

# ---------------------------------------------------------------- .bas parsing

PALETTE_RE = re.compile(
    r"Public\s+Const\s+(\w+)\s+As\s+Long\s*=\s*RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)",
    re.I,
)
SCALAR_RE = re.compile(
    r"Public\s+Const\s+(\w+)\s+As\s+(?:Single|Double|Boolean|String|Long)\s*=\s*(.+)",
    re.I,
)


def split_args(s: str) -> list[str]:
    """Split on top-level commas, ignoring commas inside parens or quotes."""
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
    if cur:
        out.append("".join(cur).strip())
    return out


def strip_vba_comment(s: str) -> str:
    """Cut a trailing VBA comment: first apostrophe outside string literals."""
    q = False
    for i, ch in enumerate(s):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return s[:i]
    return s


def untext(t: str) -> str:
    """Rebuild "A" & vbLf & "B" into real newlines."""
    t = t.strip()
    parts = re.split(r'"\s*&\s*vbLf\s*&\s*"', t, flags=re.I)
    return "\n".join(p.strip().strip('"') for p in parts).strip()


def load_bas(out_dir: str):
    """Read every .bas in out_dir.

    Returns (palette, scalars, ops, decor) where ops is the ORDERED call
    stream (bg / picture / node / path / bars already expanded into nodes)
    and decor holds the deferred SetPara / SetPartColor / BringToFront lines.
    """
    files = sorted(f for f in os.listdir(out_dir) if f.lower().endswith(".bas"))

    def key(fn: str):
        stem = os.path.splitext(fn)[0]
        m = re.search(r"(\d+)$", stem)
        return (0 if "engine" in stem.lower() else 1, int(m.group(1)) if m else 1, stem)
    files.sort(key=key)

    pal: dict[str, tuple[int, int, int]] = {}
    scal: dict[str, str] = {}
    ops: list[dict] = []
    decor: list[str] = []
    const_names = {"LINE_SOLID", "LINE_DASH", "LINE_DOT", "LINE_DASHDOT",
                   "True", "False", "Nothing", "msoFalse"}

    for fn in files:
        raw = _read_bas(os.path.join(out_dir, fn))
        # only CONTENT modules contribute call lines; the engine declares the
        # shared constants but its internal AddNode/AddPath calls (fallbacks,
        # AddBars expansion) must NOT enter the ops stream
        is_content = "content" in os.path.splitext(fn)[0].lower()
        for m in PALETTE_RE.finditer(raw):
            pal[m.group(1).lower()] = (int(m.group(2)), int(m.group(3)), int(m.group(4)))
        for m in SCALAR_RE.finditer(raw):
            scal[m.group(1).lower()] = strip_vba_comment(m.group(2)).strip()
        # join VBA line continuations so multi-line calls parse as one line
        logical = re.sub(r"_\s*\r?\n\s*", " ", raw)
        if not is_content:
            continue
        for line in logical.splitlines():
            line = strip_vba_comment(line).strip()
            if line.startswith("AddBG "):
                a = split_args(line[6:])
                ops.append(dict(op="bg", id=a[1].strip('"'),
                                c1=a[2], c2=a[3] if len(a) > 3 else "-1"))
            elif line.startswith("AddPicture "):
                a = split_args(line[11:])
                ops.append(dict(op="picture", id=a[1].strip('"'),
                                file=a[2].strip('"'),
                                x=num(a[3]), y=num(a[4]),
                                w=num(a[5]) if len(a) > 5 else 0.0,
                                h=num(a[6]) if len(a) > 6 else 0.0))
            elif line.startswith("AddNode "):
                ops.append(dict(op="node", **parse_node(split_args(line[8:])[1:])))
            elif line.startswith("AddPath "):
                ops.append(dict(op="path", **parse_path(split_args(line[8:]))))
            elif line.startswith("AddBars "):
                ops.extend(parse_bars(split_args(line[8:])))
            elif line.startswith(("SetPara ", "SetPartColor ", "BringToFront ")):
                decor.append(line)  # second pass, after ALL nodes exist
    return pal, scal, ops, decor


def num(tok: str, default: float = 0.0) -> float:
    tok = (tok or "").strip()
    m = re.match(r"^[+-]?[\d.]+", tok)
    return float(m.group(0)) if m else default


def parse_node(args: list[str]) -> dict:
    """args = [id, kind, x, y, w, h, fill, line, lw, dash, text?, pt?, bold?, ...]"""
    n = dict(
        id=args[0].strip('"'), kind=args[1].strip('"').lower(),
        x=num(args[2]), y=num(args[3]), w=num(args[4]), h=num(args[5]),
        fill=args[6] if len(args) > 6 else "-1",
        line=args[7] if len(args) > 7 else "-1",
        lw=num(args[8], 1.0) if len(args) > 8 else 1.0,
        dash=args[9] if len(args) > 9 else "LINE_SOLID",
    )
    rest = args[10:]
    n["text"] = untext(rest[0]) if rest and rest[0] else ""
    n["pt"] = num(rest[1], 11.0) if len(rest) > 1 and rest[1] else 11.0
    n["bold"] = len(rest) > 2 and rest[2].strip().lower() == "true"
    n["ital"] = len(rest) > 3 and rest[3].strip().lower() == "true"
    n["fc"] = rest[4] if len(rest) > 4 and rest[4] else "-1"
    n["corner"] = num(rest[5], -1.0) if len(rest) > 5 and rest[5] else -1.0
    n["fname"] = rest[6].strip('"') if len(rest) > 6 and rest[6] else ""
    n["adj2"] = num(rest[7], -1.0) if len(rest) > 7 and rest[7] else -1.0
    return n


def parse_path(args: list[str]) -> dict:
    """args = [sld, id, "x1,y1;x2,y2", color, lw, dash, arrowEnd]"""
    pts = []
    for pair in args[2].strip('"').split(";"):
        xy = split_args(pair)
        if len(xy) >= 2:
            pts.append((num(xy[0]), num(xy[1])))
    return dict(
        id=args[1].strip('"'), pts=pts,
        color=args[3] if len(args) > 3 else "-1",
        lw=num(args[4], 1.0) if len(args) > 4 else 1.0,
        dash=args[5] if len(args) > 5 else "LINE_SOLID",
        arrow=len(args) > 6 and args[6].strip().lower() == "true",
    )


def fmt_num(v: float) -> str:
    """Mirror VBA Format$(v, "0.##") for value labels."""
    s = ("%.2f" % v).rstrip("0").rstrip(".")
    return s if s else "0"


def parse_bars(a: list[str]) -> list[dict]:
    """Expand AddBars into node/path ops, mirroring AddBars in the engine.

    args = [sld, id, x, y, w, h, title, cats, vals, barC, lineC, txtC
            [, labelPt [, titlePt [, fontName]]]]
    """
    nid = a[1].strip('"')
    x, y, w, h = num(a[2]), num(a[3]), num(a[4]), num(a[5])
    title = untext(a[6]) if len(a) > 6 else ""
    cats = a[7].strip('"') if len(a) > 7 else ""
    vals = a[8].strip('"') if len(a) > 8 else ""
    bar_c = a[9] if len(a) > 9 else "RGB(68,114,196)"
    line_c = a[10] if len(a) > 10 else "RGB(120,120,120)"
    txt_c = a[11] if len(a) > 11 else "RGB(31,31,31)"
    label_pt = num(a[12], 9.0) if len(a) > 12 and a[12] else 9.0
    title_pt = num(a[13], 11.0) if len(a) > 13 and a[13] else 11.0
    fname = a[14].strip('"') if len(a) > 14 and a[14] else ""

    try:
        v = [float(t.strip()) for t in vals.split(",") if t.strip() != ""]
    except ValueError:
        v = []
    c = [t.strip() for t in cats.split(",")]
    n = len(v)
    if n < 1:
        return [dict(op="node", **parse_node(
            ['%s_empty' % nid, "rect", repr(x), repr(y), repr(w), repr(h),
             "-1", "-1", "1", "LINE_DASH", '"数据图占位"', "9", "False", "False",
             "RGB(120,120,120)"]))]

    vmax = max(max(v), 1.0)
    title_h = 0.16 * h if title else 0.0
    plot_top = y + title_h
    cat_h = 0.14 * h
    ybase = y + h - cat_h
    plot_h = ybase - plot_top
    slot = w / n

    out: list[dict] = []
    if title:
        out.append(dict(op="node", **parse_node(
            ['%s_t' % nid, "rect", repr(x), repr(y), repr(w), repr(title_h),
             "-1", "-1", "1", "LINE_SOLID", '"%s"' % title.replace('"', ''),
             repr(title_pt), "True", "False", txt_c, "-1", '"%s"' % fname])))
    for i in range(n):
        bx = x + slot * i + 0.2 * slot
        bh = (v[i] / vmax) * plot_h * 0.9
        bh = min(bh, plot_h)
        by = ybase - bh
        out.append(dict(op="node", **parse_node(
            ['%s_b%d' % (nid, i + 1), "rect", repr(bx), repr(by),
             repr(0.6 * slot), repr(bh), bar_c, "-1", "1", "LINE_SOLID",
             '""', "8"])))
        ly = by - 0.08 * h
        if ly < plot_top:
            ly = plot_top
        out.append(dict(op="node", **parse_node(
            ['%s_v%d' % (nid, i + 1), "rect", repr(bx - 0.1 * slot), repr(ly),
             repr(slot), repr(0.08 * h), "-1", "-1", "1", "LINE_SOLID",
             '"%s"' % fmt_num(v[i]), repr(label_pt), "False", "False",
             txt_c, "-1", '"%s"' % fname])))
        if i < len(c) and c[i]:
            out.append(dict(op="node", **parse_node(
                ['%s_c%d' % (nid, i + 1), "rect", repr(x + slot * i), repr(ybase),
                 repr(slot), repr(cat_h), "-1", "-1", "1", "LINE_SOLID",
                 '"%s"' % c[i], repr(label_pt), "False", "False",
                 txt_c, "-1", '"%s"' % fname])))
    out.append(dict(op="path", **parse_path(
        ["sld", '%s_axis' % nid,
         '"%s,%s;%s,%s"' % (repr(x), repr(ybase), repr(x + w), repr(ybase)),
         line_c, "1", "LINE_SOLID", "False"])))
    return out


# ---------------------------------------------------------------- geometry

def resolve_color(tok: str, pal: dict):
    tok = (tok or "").strip()
    low = tok.lower()
    if low in pal:
        return pal[low]
    m = re.match(r"RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", tok, re.I)
    if m:
        return tuple(int(g) for g in m.groups())
    if tok in ("-1", ""):
        return None
    return (0, 0, 0)


def resolve_scalar(tok: str, scal: dict, default: float) -> float:
    """Resolve a numeric literal that may reference another Public Const."""
    tok = (tok or "").strip()
    if not tok:
        return default
    if tok.lower() in scal:
        return resolve_scalar(scal[tok.lower()], scal, default)
    return num(tok, default)


def canvas_geometry(scal: dict):
    """Reproduce the geometry constants the engine relies on."""
    cw = resolve_scalar(scal.get("canvas_w_px", "1200"), scal, 1200.0)
    ch = resolve_scalar(scal.get("canvas_h_px", "675"), scal, 675.0)
    slide_w_in = resolve_scalar(scal.get("slide_w_in", "13.333"), scal, 13.333)
    slide_h_in = resolve_scalar(scal.get("slide_h_in", "7.5"), scal, 7.5)
    custom = scal.get("custom_size", "False").strip().lower() == "true"
    margin = 0.35
    avail_w = slide_w_in - 2 * margin
    avail_h = slide_h_in - 2 * margin
    px_to_in = resolve_scalar(scal.get("px_to_in", ""), scal, 0.0)
    if px_to_in <= 0:
        px_to_in = min(avail_w / cw, avail_h / ch)
    off_x = resolve_scalar(scal.get("offset_x_in", ""), scal, -1.0)
    off_y = resolve_scalar(scal.get("offset_y_in", ""), scal, -1.0)
    if off_x < 0:
        off_x = margin + (avail_w - cw * px_to_in) / 2
    if off_y < 0:
        off_y = margin + (avail_h - ch * px_to_in) / 2
    font_name = scal.get("font_name", '"Arial"').strip().strip('"') or "Arial"
    font_cn = scal.get("font_name_cn", "").strip().strip('"') or font_name
    asset_dir = scal.get("asset_dir", '""').strip().strip('"')
    return dict(cw=cw, ch=ch, sw=slide_w_in, sh=slide_h_in, custom=custom,
                s=px_to_in, ox=off_x, oy=off_y, font=font_name,
                font_cn=font_cn, assets=asset_dir)


# ---------------------------------------------------------------- path A: COM

ENGINE_NAME = "modPoster_Engine"
CONTENT_PREFIX = "modPoster_Content"


def is_powerpoint_available() -> bool:
    if os.name != "nt":
        return False
    # do not let gen_py cache shadow the real typelib
    import importlib
    try:
        mod = importlib.import_module("win32com.client")
    except ImportError:
        return False
    try:
        mod.gencache.EnsureDispatch("PowerPoint.Application")
        return True
    except Exception:
        return False


def _read_bas(path):
    """Read a .bas file trying UTF-8 then GBK (zh-CN VBE), then latin-1."""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("latin-1")


def module_name(text: str, fallback: str) -> str:
    m = re.search(r"Attribute\s+VB_Name\s*=\s*\"([^\"]+)\"", text)
    return m.group(1) if m else fallback


def build_with_com(out_dir: str, out_path: str, bas_files: list[str],
                   run: bool, export_png: str | None) -> int:
    import win32com.client as w32

    app = w32.gencache.EnsureDispatch("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Add()
    # start from a blank 16:9 board; BuildPoster resizes when CUSTOM_SIZE
    pres.PageSetup.SlideWidth = 13.333 * 72
    pres.PageSetup.SlideHeight = 7.5 * 72

    # order matters: engine first, then content modules by trailing index
    def key(p: str):
        stem = os.path.splitext(os.path.basename(p))[0]
        m = re.search(r"(\d+)$", stem)
        return (0 if "engine" in stem.lower() else 1, int(m.group(1)) if m else 1, stem)
    bas_files = sorted(bas_files, key=key)

    vbproj = pres.VBProject
    for p in bas_files:
        text = bas_text(p)
        name = module_name(text, os.path.splitext(os.path.basename(p))[0])
        comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = name
        # strip the Attribute line: a component we just created already carries it
        body = re.sub(r"^Attribute\s+VB_Name\s*=.*\r?\n", "", text, count=1)
        comp.CodeModule.AddFromString(body)
        print("  COM: imported %s" % name)

    if not run:
        pres.SaveAs(out_path)
        print("  COM: saved template (BuildPoster NOT executed)")
        return 0

    try:
        app.Run("BuildPoster")
    except Exception as exc:  # surface the VBA error instead of hiding it
        print("  COM: BuildPoster raised -> %s" % exc, file=sys.stderr)
        pres.SaveAs(out_path)
        return 2

    if export_png and pres.Slides.Count >= 1:
        try:
            sw, sh = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
            w_px = 1920
            h_px = int(round(1920 * sh / sw))
            pres.Slides(1).Export(export_png, "PNG", w_px, h_px)
            print("  COM: exported preview %s" % export_png)
        except Exception as exc:
            print("  COM: PNG export failed -> %s" % exc, file=sys.stderr)

    pptm = os.path.splitext(out_path)[0] + ".pptm"
    try:
        pres.SaveAs(pptm)
    except Exception as exc:
        print("  COM: SaveAs .pptm failed -> %s" % exc, file=sys.stderr)
    try:
        if os.path.exists(pptm):
            tmp = app.Presentations.Open(pptm, WithWindow=False)
            tmp.SaveAs(out_path, 24)  # ppSaveAsOpenXMLPresentation
            tmp.Close()
    except Exception as exc:
        print("  COM: .pptx conversion failed -> %s" % exc, file=sys.stderr)
    return 0


# ---------------------------------------------------------------- path B: replay

# name -> msoAutoShapeType, mirrors ShapeTypeFor() in modPoster_Engine.bas
KIND_TO_Mso = {
    "rect": 1, "rectangle": 1, "box": 1,
    "round_rect": 5, "rounded": 5, "rounded_rect": 5,
    "oval": 9, "ellipse": 9, "circle": 9,
    "diamond": 4, "decision": 4,
    "parallelogram": 2,
    "trapezoid": 3,
    "pentagon": 51,
    "pentagon_reg": 12,
    "hexagon": 10,
    "stadium": 5, "pill": 5, "terminator": 5,
    "can": 13, "cylinder": 13,
    "document": 61,
    "right_arrow": 33, "arrow_right": 33,
    "left_arrow": 34, "arrow_left": 34,
    "up_arrow": 35, "arrow_up": 35,
    "down_arrow": 36, "arrow_down": 36,
    "chevron": 52,
    "arc": 25,
    "star4": 91, "star5": 92, "star8": 93, "star6": 147,
    "heart": 21, "cloud": 179, "sun": 23, "moon": 24,
    "lightning": 22, "bolt": 22, "smiley": 17, "donut": 18,
    "wave": 103, "double_wave": 104,
    "callout": 106, "round_callout": 106,
    "bubble": 107, "oval_callout": 107,
}


def resolve_asset(fname: str, out_dir: str, asset_dir: str):
    """Try fname as-is, then against ASSET_DIR, then the output dir."""
    cands = []
    adir = asset_dir
    if adir and not os.path.isabs(adir):
        adir = os.path.join(out_dir, adir)
    if os.path.isabs(fname):
        cands.append(fname)
        if adir:
            cands.append(os.path.join(adir, os.path.basename(fname)))
    else:
        if adir:
            cands.append(os.path.join(adir, fname))
        cands.append(os.path.join(out_dir, fname))
        cands.append(os.path.join(out_dir, "assets", fname))
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def _set_run_font(r, name: str, ea_name: str, size_pt: float, bold: bool,
                  italic: bool, color):
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    r.font.size = Pt(size_pt)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = name
    r.font.color.rgb = RGBColor(*color)
    # East-Asian font: a:ea (+ a:cs) right after a:latin in rPr
    rPr = r._r.get_or_add_rPr()
    from pptx.oxml.ns import qn
    latin = rPr.find(qn("a:latin"))
    prev = latin
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            if prev is not None:
                prev.addnext(el)
            else:
                rPr.append(el)
        el.set("typeface", ea_name)
        prev = el


def build_with_replay(out_dir: str, out_path: str) -> int:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    pal, scal, ops, decor = load_bas(out_dir)
    g = canvas_geometry(scal)
    counts = {}
    for o in ops:
        counts[o["op"]] = counts.get(o["op"], 0) + 1
    print("  replay: ops -> " + ", ".join("%s %d" % (k, v)
                                          for k, v in sorted(counts.items())))

    prs = Presentation()
    prs.slide_width = Inches(g["sw"])
    prs.slide_height = Inches(g["sh"])
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    def X(px):
        return Inches(g["ox"] + px * g["s"])

    def Y(py):
        return Inches(g["oy"] + py * g["s"])

    def L(v):
        return Inches(v * g["s"])

    nodes_by_id: dict[str, dict] = {}

    def add_node_shp(n: dict):
        nodes_by_id[n["id"]] = n
        mso_id = KIND_TO_Mso.get(n["kind"], 1)
        try:
            shape_type = MSO_SHAPE(mso_id)
        except ValueError:
            shape_type = MSO_SHAPE.RECTANGLE
        try:
            shp = slide.shapes.add_shape(
                shape_type, X(n["x"]), Y(n["y"]), L(n["w"]), L(n["h"]))
        except Exception:
            shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, X(n["x"]), Y(n["y"]), L(n["w"]), L(n["h"]))

        fill = resolve_color(n["fill"], pal)
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(*fill)

        linec = resolve_color(n["line"], pal)
        if linec is None:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = RGBColor(*linec)
            shp.line.width = Pt(max(0.5, n["lw"]))
            if n["dash"].upper() in ("LINE_DASH", "LINE_DASHDOT", "LINE_DOT"):
                from pptx.enum.dml import MSO_LINE_DASH_STYLE
                shp.line.dash_style = (
                    MSO_LINE_DASH_STYLE.DASH if n["dash"].upper() == "LINE_DASH"
                    else MSO_LINE_DASH_STYLE.ROUND_DOT if n["dash"].upper() == "LINE_DOT"
                    else MSO_LINE_DASH_STYLE.DASH_DOT)

        corner = 0.5 if n["kind"] in ("stadium", "pill", "terminator") else n["corner"]
        if corner and corner > 0:
            try:
                shp.adjustments[0] = max(0.0, min(0.5, corner))
            except Exception:
                pass
        if n["adj2"] and n["adj2"] > 0:
            try:
                shp.adjustments[1] = max(0.0, min(1.0, n["adj2"]))
            except Exception:
                pass

        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Pt(1)
        tf.margin_top = tf.margin_bottom = Pt(0)
        if n["text"]:
            tf.text = n["text"]
            tc = resolve_color(n["fc"], pal) or (31, 31, 31)
            latin = n.get("fname") or g["font"]
            ea = n.get("fname") or g["font_cn"]
            align = n.get("align") or "center"
            ml_px = n.get("ml", 0) or 0
            for para in tf.paragraphs:
                para.alignment = (PP_ALIGN.LEFT if align == "left"
                                  else PP_ALIGN.RIGHT if align == "right"
                                  else PP_ALIGN.CENTER)
                for r in para.runs:
                    _set_run_font(r, latin, ea, n["pt"], n["bold"], n["ital"], tc)
            if ml_px and align == "left":
                tf.margin_left = Pt(ml_px * g["s"] * 72)
            prefix = n.get("prefix")
            if prefix and prefix[0] > 0:
                nc, ctoken = prefix
                pc = resolve_color(ctoken, pal) or (200, 30, 40)
                p0 = tf.paragraphs[0]
                full = p0.runs[0].text if p0.runs else ""
                if full:
                    p0.clear()
                    r1 = p0.add_run()
                    r1.text = full[:nc]
                    r2 = p0.add_run()
                    r2.text = full[nc:]
                    _set_run_font(r1, latin, ea, n["pt"], True, n["ital"], pc)
                    _set_run_font(r2, latin, ea, n["pt"], n["bold"], n["ital"], tc)

    def add_path_shp(p: dict):
        pts = p["pts"]
        c = resolve_color(p["color"], pal) or (0, 0, 0)
        for i in range(len(pts) - 1):
            (x0, y0), (x1, y1) = pts[i], pts[i + 1]
            conn = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, X(x0), Y(y0), X(x1), Y(y1))
            conn.line.color.rgb = RGBColor(*c)
            conn.line.width = Pt(max(0.5, p["lw"]))
            dashv = (p.get("dash") or "LINE_SOLID").upper()
            if dashv in ("LINE_DASH", "LINE_DASHDOT", "LINE_DOT"):
                from pptx.enum.dml import MSO_LINE_DASH_STYLE
                conn.line.dash_style = (
                    MSO_LINE_DASH_STYLE.DASH if dashv == "LINE_DASH"
                    else MSO_LINE_DASH_STYLE.ROUND_DOT if dashv == "LINE_DOT"
                    else MSO_LINE_DASH_STYLE.DASH_DOT)
            if p["arrow"] and i == len(pts) - 2:
                ln = conn.line._get_or_add_ln()
                from pptx.oxml.ns import qn
                tail = ln.makeelement(qn("a:tailEnd"),
                                      {"type": "triangle", "w": "med", "len": "med"})
                ln.append(tail)

    def add_bg_shp(o: dict):
        c1 = resolve_color(o["c1"], pal) or (255, 255, 255)
        c2 = resolve_color(o["c2"], pal)
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                     Inches(g["sw"]), Inches(g["sh"]))
        shp.line.fill.background()
        if c2 is None:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(*c1)
        else:
            shp.fill.gradient()
            stops = shp.fill.gradient_stops
            stops[0].color.rgb = RGBColor(*c1)
            stops[0].position = 0.0
            stops[1].color.rgb = RGBColor(*c2)
            stops[1].position = 1.0
            # set the linear-gradient angle directly in OOXML: 5400000 =
            # 90 deg clockwise -> vector points down -> c1 top, c2 bottom
            # (python-pptx gradient_angle uses opposite semantics)
            from pptx.oxml.ns import qn
            gf = shp._element.spPr.find(qn("a:gradFill"))
            if gf is not None:
                lin = gf.find(qn("a:lin"))
                if lin is None:
                    lin = gf.makeelement(qn("a:lin"), {})
                    gf.append(lin)
                lin.set("ang", "5400000")
                lin.set("scaled", "0")

    def add_picture_shp(o: dict):
        resolved = resolve_asset(o["file"], out_dir, g["assets"])
        if resolved is None:
            # parity with the VBA missing-asset fallback
            pw = o["w"] if o["w"] > 0 else 120.0
            ph = o["h"] if o["h"] > 0 else 90.0
            add_node_shp(parse_node([
                o["id"] + "_ph", "rect", repr(o["x"]), repr(o["y"]),
                repr(pw), repr(ph), "RGB(238,238,238)", "RGB(170,170,170)",
                "1", "LINE_DASH", '"IMG %s"' % o["file"], "9", "False", "False",
                "RGB(120,120,120)"]))
            print("  replay: missing asset %r -> placeholder box" % o["file"])
            return
        try:
            slide.shapes.add_picture(resolved, X(o["x"]), Y(o["y"]),
                                     L(o["w"]), L(o["h"]))
        except Exception:
            # natural-size mode (w/h <= 0) or a broken file
            try:
                slide.shapes.add_picture(resolved, X(o["x"]), Y(o["y"]))
            except Exception as exc:
                print("  replay: add_picture failed for %r -> %s"
                      % (o["file"], exc), file=sys.stderr)

    # ---- deferred decor: patch node dicts BEFORE drawing so that align,
    # margins and prefix colors take effect in add_node_shp ----
    for o in ops:
        if o["op"] == "node":
            nodes_by_id.setdefault(o["id"], o)
    for line in decor:
        if line.startswith("SetPara "):
            args = split_args(line[8:])
            nid, align = args[1].strip('"'), args[2].strip('"').lower()
            ml = num(args[3], 0.0) if len(args) > 3 else 0.0
            n = nodes_by_id.get(nid)
            if n:
                n["align"], n["ml"] = align, ml
        elif line.startswith("SetPartColor "):
            args = split_args(line[8:])
            nid = args[1].strip('"')
            nc = int(num(args[2], 0.0)) if len(args) > 2 else 0
            ctoken = args[3] if len(args) > 3 else "RGB(200, 30, 40)"
            n = nodes_by_id.get(nid)
            if n:
                n["prefix"] = (nc, ctoken)
        elif line.startswith("BringToFront "):
            nid = split_args(line[13:])[0].strip('"')
            if nid not in nodes_by_id:
                print("  replay: BringToFront target %r not found" % nid, file=sys.stderr)
            # replay keeps VBA call order, so BringToFront is accepted as a no-op
            print("  replay: note BringToFront %r is a no-op (call order already on top)" % nid)

    # ---- draw in call order (z-order parity with the VBA path) ----
    for o in ops:
        if o["op"] == "bg":
            add_bg_shp(o)
        elif o["op"] == "picture":
            add_picture_shp(o)
        elif o["op"] == "node":
            add_node_shp(o)
        elif o["op"] == "path":
            add_path_shp(o)

    prs.save(out_path)
    return 0


# ---------------------------------------------------------------- driver

def slide_shape_count(pptx_path: str) -> int:
    try:
        from pptx import Presentation
    except ImportError:
        return -1
    prs = Presentation(pptx_path)
    return sum(len(s.shapes) for s in prs.slides)


def main() -> int:
    ap = argparse.ArgumentParser(description="poster-vba: .bas -> .pptx")
    ap.add_argument("out_dir", help="directory holding modPoster_Engine.bas + modPoster_Content*.bas")
    ap.add_argument("--out", default="poster.pptx", help="output pptx file name")
    ap.add_argument("--no-run", action="store_true",
                    help="COM path only: build the file without executing BuildPoster")
    ap.add_argument("--replay", action="store_true",
                    help="force the no-Office replay path even if PowerPoint is installed")
    ap.add_argument("--export-png", default=None, metavar="PNG",
                    help="COM path: export slide 1 as PNG for the visual review")
    args = ap.parse_args()

    out_dir = os.path.abspath(args.out_dir)
    if not os.path.isdir(out_dir):
        print("output dir not found: %s" % out_dir, file=sys.stderr)
        return 2
    bas_files = [os.path.join(out_dir, f) for f in sorted(os.listdir(out_dir))
                 if f.lower().endswith(".bas")]
    if not bas_files:
        print("no .bas files in %s" % out_dir, file=sys.stderr)
        return 2
    engine = [p for p in bas_files if "engine" in os.path.basename(p).lower()]
    content = [p for p in bas_files if "content" in os.path.basename(p).lower()]
    if not engine:
        print("warning: modPoster_Engine.bas missing from %s" % out_dir, file=sys.stderr)
    if not content:
        print("warning: no modPoster_Content*.bas found in %s" % out_dir, file=sys.stderr)

    out_path = args.out if os.path.isabs(args.out) else os.path.join(out_dir, args.out)
    print("poster-vba :: %s -> %s" % (out_dir, out_path))

    used = None
    if not args.replay and is_powerpoint_available():
        print("[path A] PowerPoint COM automation")
        try:
            rc = build_with_com(out_dir, out_path, bas_files,
                                run=not args.no_run, export_png=args.export_png)
            used = "COM"
            if rc != 0 and not os.path.exists(out_path):
                used = None
        except Exception as exc:
            print("  COM path failed -> %s" % exc, file=sys.stderr)
            used = None

    if used is None:
        print("[path B] .bas replay via python-pptx (no PowerPoint needed)")
        try:
            build_with_replay(out_dir, out_path)
            used = "replay"
        except ImportError as exc:
            print("python-pptx is required for the replay path: %s" % exc, file=sys.stderr)
            return 2

    if not os.path.exists(out_path):
        print("FAILED: %s was not written" % out_path, file=sys.stderr)
        return 1
    count = slide_shape_count(out_path)
    size_kb = os.path.getsize(out_path) / 1024.0
    print("OK  [%s]  %s  (%.1f KB, %d shapes)" % (used, out_path, size_kb, count))
    if count == 0:
        print("FAILED: the slide has no shapes -- the poster was not drawn", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
