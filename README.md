# poster-vba

[🇬🇧 English](#english) · [🇨🇳 中文](#zh)

---

<h2 id="english">🇬🇧 English</h2>

A WorkBuddy skill that turns competition posters, exhibition boards and playful
cartoon-style diagrams into **editable PowerPoint files** — a real `.pptx` plus
the VBA source that can rebuild it at any time.

### What it does

Given a poster image, the skill produces VBA code that draws native, editable
PowerPoint shapes (titles, panels, body text, gradient background, cartoon
shapes) and embeds icons as **transparent PNG pictures**. Open the generated
`.pptx` and everything is already drawn and editable — no picture-in-a-picture.

### Sibling of sci-flowchart-vba

Same lineage, different direction: that one handles academic flowcharts, this one
handles posters. Four areas where posters need more than a flowchart engine:

| Need | How this skill handles it |
|---|---|
| **Cartoon icons** | Three routes: ① search a transparent PNG online; ② `scripts/icon_tool.py` cuts + sharpens from the source image; ③ engine shape collage as last resort. Embedded with `AddPicture`; a missing file degrades to a placeholder box instead of failing. |
| **Chinese fonts** | Dual font slots: `FONT_NAME` (Latin) + `FONT_NAME_CN` (Chinese, written to the `NameFarEast` slot). `references/font-matching.md` recognizes the family from glyph features, `scripts/font_check.py` verifies it is actually installed. Point sizes are measured from pixels — never a blanket 12pt. |
| **Data charts** | Ask the user first: leave blank → dashed placeholder box; draw a simple one → `AddBars` native-shape mini bar chart (no Excel dependency, and both build paths stay pixel-identical). |
| **Background** | `AddBG` solid or top→bottom gradient; when decoration is complex, `icon_tool.py` cuts a clean region and pastes it back; when the source text is burned into the background, an opaque panel cover strategy hides it. |

### Two modes

- **Mode 1 — replicate**: one poster image plus "replicate / redraw / make it
  editable". Programmatic color and font measurement + multimodal vision rebuild
  the layout, assets and text.
- **Mode 2 — style creation**: a reference poster (for style) + a content list or
  paper paragraph (for content). The palette, font and layout language come from
  the reference; the content is new.

### Layout

```
poster-vba/
  SKILL.md                  # skill manifest + full contract for the code model
  assets/
    modPoster_Engine.bas    # fixed engine (do not rewrite)
  references/
    example_content.bas     # runnable content-module example
    poster-module-map.md    # engine API table, kind -> mso ids, path parity rules
    font-matching.md        # Chinese font feature recognition + point-size formulas
    icon-sourcing.md        # icon sourcing routes (search / cut / collage) + BG strategy
  scripts/
    build_poster.py         # final step: .bas -> real .pptx (COM, replay fallback)
    icon_tool.py            # probe / check / cut (de-background, trim, upscale, sharpen)
    font_check.py           # enumerate installed fonts + match/fallback
    measure_colors.py       # generic color sampling (dominant / region / text / accent)
    compare_text.py         # text-band point-size calibration (review round 1)
```

### Quick start (as a user)

1. Get the `.pptx` — the skill already generates it via `build_poster.py`.
2. Or rebuild yourself: in PowerPoint press `Alt+F11`, right-click the project,
   **Import File** for each `.bas` (engine + content), then `F5` → `BuildPoster`.
3. Run `RemovePoster` to clear and redraw.

### Contract highlights (for the code model)

- The content module must declare **11 geometry constants** (including
  `FONT_NAME_CN` and `ASSET_DIR`) and one `Public Sub DrawAll(sld As Slide)`,
  whose first call is the background subroutine.
- Colors are `RGB(r,g,b)` literals; no `vbRed`; no empty argument slots (`,,`).
- `.bas` files are **UTF-8 encoded + CRLF** — keep the source language, so a
  Chinese poster keeps Chinese text. That is what lets Chinese survive into the
  `.pptx`. (On a Chinese Windows, save as ANSI/GBK before *manual* import to
  avoid VBE mojibake; not needed when using `build_poster.py`.)
- **z-order = call order**: background first, icons drawn after the panel they
  belong to so they sit on top.
- Two-line text must be explicit: `"A" & vbLf & "B"` — never rely on autofit.
- Split into `modPoster_Content2.bas` when nodes + pictures > 18 or a single
  module > 300 lines, and call every `DrawContentN` from `DrawAll`.
- **Two review rounds are mandatory before delivery**: round 1 is a programmatic
  comparison against the source figure (palette constants vs. `measure_colors.py`
  output, fonts verified by `font_check.py`, hierarchical point sizes); round 2 is
  a visual comparison of the rendered `preview.png` against the original
  (layout, background, icon clarity, no text overflow, no tofu glyphs).

### License

MIT — do whatever you like; attribution appreciated.

---

<h2 id="zh">🇨🇳 中文</h2>

一个 WorkBuddy 技能，把**比赛展板、海报、略活泼的卡通示意图**转换成
**可编辑的 PowerPoint 文件**——一份真正能打开的 `.pptx`，外加可随时重建它的
VBA 源码。

### 它做什么

给定一张展板图片，技能生成绘制「原生、可编辑」PowerPoint 形状的 VBA 代码
（标题、面板、正文、渐变背景、卡通形状），图标以**透明 PNG 图片**嵌入。
直接打开生成好的 `.pptx`，形状已经画好且全部可编辑——不是一整张「图套图」。

### 与 sci-flowchart-vba 的关系

同源不同向：那一个做学术流程图，这一个做展板。相比流程图引擎，展板有四项
额外需求：

| 需求 | 本技能的做法 |
|---|---|
| **卡通图标** | 三条路线：① 联网搜透明 PNG；② `scripts/icon_tool.py` 从原图抠图 + 清晰化；③ 引擎形状拼贴保底。用 `AddPicture` 嵌入，素材缺失时自动降级为占位框而不是失败。 |
| **中文字体** | 双字体槽：`FONT_NAME`（西文）+ `FONT_NAME_CN`（中文，写入 `NameFarEast` 槽）。`references/font-matching.md` 按字形特征识别家族，`scripts/font_check.py` 验证字体确实已安装。字号按字像素高测算，禁止一律 12pt。 |
| **数据图** | 先问用户偏好：留空 → 虚线占位框；简单画一个 → `AddBars` 原生形状迷你柱状图（不依赖 Excel，两条构建路径像素级一致）。 |
| **背景** | `AddBG` 纯色或上→下渐变；装饰复杂时 `icon_tool.py` 抠干净区域贴回；原文字被烧进背景时用不透明面板盖板策略遮住。 |

### 两种模式

- **模式① 复刻**：一张展板图 + 「复刻 / 重绘 / 转成可编辑」。程序化取色取字号
  + 多模态识图，重建布局、素材与文字。
- **模式② 风格创作**：参考展板图（提供风格）+ 内容清单或论文段落（提供内容）。
  配色、字体、版式语言取自参考图，内容全部换新。

### 目录结构

```
poster-vba/
  SKILL.md                  # 技能清单 + 给代码模型的完整契约
  assets/
    modPoster_Engine.bas    # 固定引擎（勿改写）
  references/
    example_content.bas     # 可运行的内容模块示例
    poster-module-map.md    # 引擎 API 全表、kind→mso 编号、双路径一致性约定
    font-matching.md        # 中文字体特征识别 + 字号测算公式
    icon-sourcing.md        # 图标三路线（搜索/抠图/形状拼贴）+ 背景策略表
  scripts/
    build_poster.py         # 最后一步：.bas -> 真实 .pptx（COM 优先，回放兜底）
    icon_tool.py            # probe / check / cut（抠底、去白边、放大、锐化）
    font_check.py           # 已安装字体枚举与匹配回退
    measure_colors.py       # 通用取色（主色/区域底色/文字色/强调色）
    compare_text.py         # 文字带字号校准（评审一）
```

### 快速开始（作为使用者）

1. 拿到 `.pptx`——技能已通过 `build_poster.py` 生成。
2. 或自己重建：PowerPoint 里按 `Alt+F11`，右键工程 → **Import File** 导入每个
   `.bas`（引擎 + 内容），然后 `F5` → `BuildPoster`。
3. 运行 `RemovePoster` 可清空重画。

### 出码契约要点（给代码模型）

- 内容模块必须声明 **11 个几何常量**（含 `FONT_NAME_CN`、`ASSET_DIR`），并定义
  唯一的 `Public Sub DrawAll(sld As Slide)`，其中**最先调用背景子过程**。
- 颜色用 `RGB(r,g,b)` 字面量；不要 `vbRed` 之类；不许留空实参槽位（`,,`）。
- `.bas` 必须**UTF-8 编码 + CRLF**——文字保留源语言，原图是中文就直接写中文，
  UTF-8 才能让中文正常落地到 `.pptx`。（中文 Windows 下*手动*导入前另存为
  ANSI/GBK 可避免 VBE 乱码；用 `build_poster.py` 出图则无需此步。）
- **z-order = 调用顺序**：背景最先，图标画在所属面板之后（从而位于其上层）。
- 两行文字必须显式写 `"A" & vbLf & "B"`，不要依赖自动折行。
- 节点 + 图片 > 18 或单模块 > 300 行时，拆到 `modPoster_Content2.bas`，并在
  `DrawAll` 中依次调用每个 `DrawContentN`。
- **交付前必须过两轮评审**：评审一做与源图的程序化对照（调色板常量 vs.
  `measure_colors.py` 输出、`font_check.py` 验证字体、字号层级比例）；评审二
  做最终渲染 `preview.png` 与源图的视觉对照（版式、背景、图标清晰度、无文字
  溢出、无豆腐块）。

### 许可证

MIT——可随意使用，注明出处更佳。
