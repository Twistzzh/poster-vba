# -*- coding: utf-8 -*-
"""poster-vba 字体检查：枚举 Windows 已安装字体，对目标字体做匹配与回退建议。

用法:
    python font_check.py 微软雅黑 思源黑体 楷体 ...   # 检查这些字体是否可用
    python font_check.py --list                      # 列出全部含中文名的字体
    python font_check.py --list 圆                    # 按关键字过滤

数据源: 注册表(HKLM/HKCU Fonts) + C:\\Windows\\Fonts + 用户字体目录。
匹配: 精确(大小写/空格不敏感) -> 内置别名表 -> difflib 相似度 Top3。
输出每个查询: 可用 / 不可用 + 建议替代(已安装) + 下载提示(免费可获取)。
"""
import difflib
import os
import re
import sys

# 查询词 -> 常见家族名(英文/中文别名)，命中任意即视为该字体可用
ALIAS = {
    "微软雅黑": ["Microsoft YaHei", "微软雅黑"],
    "雅黑": ["Microsoft YaHei", "微软雅黑"],
    "yahei": ["Microsoft YaHei", "微软雅黑"],
    "黑体": ["SimHei", "黑体"],
    "simhei": ["SimHei", "黑体"],
    "宋体": ["SimSun", "宋体", "NSimSun"],
    "simsun": ["SimSun", "宋体"],
    "中易宋体": ["SimSun", "宋体"],
    "新宋体": ["NSimSun", "新宋体"],
    "仿宋": ["FangSong", "仿宋", "仿宋_GB2312"],
    "fangsong": ["FangSong", "仿宋"],
    "楷体": ["KaiTi", "楷体", "楷体_GB2312"],
    "kaiti": ["KaiTi", "楷体"],
    "隶书": ["LiSu", "隶书"],
    "幼圆": ["YouYuan", "幼圆"],
    "华文黑体": ["STHeiti", "华文黑体"],
    "华文细黑": ["STXihei", "华文细黑"],
    "华文楷体": ["STKaiti", "华文楷体"],
    "华文宋体": ["STSong", "华文宋体"],
    "华文中宋": ["STZhongsong", "华文中宋"],
    "华文仿宋": ["STFangsong", "华文仿宋"],
    "华文彩云": ["STCaiyun", "华文彩云"],
    "华文琥珀": ["STHupo", "华文琥珀"],
    "华文隶书": ["STLiti", "华文隶书"],
    "华文行楷": ["STXingkai", "华文行楷"],
    "华文新魏": ["STXinwei", "华文新魏"],
    "方正舒体": ["FZShuTi", "方正舒体"],
    "方正姚体": ["FZYaoTi", "方正姚体"],
    "思源黑体": ["Source Han Sans SC", "Source Han Sans CN", "Noto Sans CJK SC",
                 "Noto Sans SC", "思源黑体"],
    "思源宋体": ["Source Han Serif SC", "Source Han Serif CN", "Noto Serif CJK SC",
                 "Noto Serif SC", "思源宋体"],
    "霞鹜文楷": ["LXGW WenKai", "霞鹜文楷", "LXGW WenKai GB"],
    "站酷快乐体": ["ZCOOL KuaiLe", "站酷快乐体"],
    "站酷高端黑": ["ZCOOL QingKe HuangYou", "站酷高端黑"],
    "阿里巴巴普惠体": ["Alibaba PuHuiTi", "阿里巴巴普惠体"],
    "arial": ["Arial"],
    "times": ["Times New Roman"],
    "calibri": ["Calibri"],
    "segoe": ["Segoe UI"],
    "impact": ["Impact"],
    "comic": ["Comic Sans MS"],
}

FREE_DOWNLOAD = {
    "思源黑体": "GitHub adobe-fonts/source-han-sans (免费可商用)",
    "思源宋体": "GitHub adobe-fonts/source-han-serif (免费可商用)",
    "霞鹜文楷": "GitHub lxgw/LxgwWenKai (免费可商用)",
    "站酷快乐体": "站酷/各字体站 (免费可商用)",
    "阿里巴巴普惠体": "Alibaba Fonts 官网 (免费可商用)",
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").lower())


def _installed() -> list[str]:
    """Return family alias strings from registry value names + font dirs."""
    found: set[str] = set()
    try:
        import winreg
        for hive, path in ((winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                           (winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Fonts"),
                           (winreg.HKEY_CURRENT_USER,
                            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")):
            try:
                k = winreg.OpenKey(hive, path)
            except OSError:
                continue
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(k, i)
                except OSError:
                    break
                i += 1
                base = re.sub(r"\s*\((?:TrueType|OpenType|TrueTypeOutline|"
                              r"PostScript|CFF|CFF & OpenType)[^)]*\)\s*$", "", name)
                for part in base.split("&"):
                    part = part.strip()
                    if part:
                        found.add(part)
    except ImportError:
        pass
    for d in (r"C:\Windows\Fonts",
              os.path.join(os.environ.get("LOCALAPPDATA", ""),
                           "Microsoft", "Windows", "Fonts")):
        if os.path.isdir(d):
            for f in os.listdir(d):
                stem = os.path.splitext(f)[0]
                found.add(stem)
    return sorted(found)


def has_cjk(s: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in s)


def check_one(query: str, installed: list[str]) -> None:
    norm_q = _norm(query)
    # 1) direct substring / equality
    exact = [f for f in installed
             if _norm(f) == norm_q or _norm(f).replace("-", "") == norm_q
             or norm_q in _norm(f)]
    # 2) alias table expansion
    alias_hits: list[str] = []
    for key, fams in ALIAS.items():
        if _norm(key) == norm_q or norm_q in _norm(key) or _norm(key) in norm_q:
            for fam in fams:
                alias_hits += [f for f in installed if _norm(fam) in _norm(f)]
    hits = sorted(set(exact) | set(alias_hits))
    if hits:
        print("[可用] %s" % query)
        for h in hits[:6]:
            print("        %s" % h)
        return
    # 3) fuzzy suggestions
    close = difflib.get_close_matches(query, installed, n=3, cutoff=0.4)
    also = []
    for key, fams in ALIAS.items():
        for fam in fams:
            for f in installed:
                if difflib.SequenceMatcher(None, _norm(fam), _norm(f)).ratio() > 0.75:
                    also.append(f)
    sugg = sorted(set(close) | set(also))[:5]
    print("[未安装] %s" % query)
    if query in FREE_DOWNLOAD:
        print("        免费获取: %s" % FREE_DOWNLOAD[query])
    if sugg:
        print("        已安装的相近字体: %s" % " | ".join(sugg))


def main() -> int:
    args = sys.argv[1:]
    installed = _installed()
    if not installed:
        print("未枚举到任何已安装字体（非 Windows？）", file=sys.stderr)
        return 2
    if args and args[0] == "--list":
        kw = _norm(args[1]) if len(args) > 1 else None
        for f in installed:
            if kw is None:
                if has_cjk(f):
                    print(f)
            elif kw in _norm(f):
                print(f)
        return 0
    if not args:
        print(__doc__)
        return 2
    print("已安装字体 %d 个\n" % len(installed))
    for q in args:
        check_one(q, installed)
        print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
