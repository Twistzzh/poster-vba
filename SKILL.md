---
name: poster-vba
description: 把比赛展板/海报/略活泼的卡通示意图复刻为可编辑 PowerPoint（.pptx + VBA .bas 源码）。当用户提供一张展板、海报、宣传图、卡通风格示意图并要求"复刻/重绘/转成可编辑 PPT"，或提供参考展板图 + 文字内容要求按其风格制作新展板，或明确提到"展板转 PPT""海报复刻""比赛展板重绘""卡通示意图转 PPT""把这张图上的图标抠出来放进 PPT"时，必须使用本 skill。特色：卡通图标素材（联网搜透明 PNG 或从原图抠图+清晰化）、中文字体精准匹配（NameFarEast 槽 + 字号测算）、数据图占位或原生形状迷你柱状图、背景整版还原（渐变/抠图/盖板策略）。最终交付**已画好形状、可直接编辑的 .pptx**，附带分模块 .bas（导入 PowerPoint 运行 BuildPoster 可随时重建）。
agent_created: true
---

# 比赛展板/卡通示意图 → PowerPoint VBA（poster-vba）

把比赛展板、海报、略活泼的卡通示意图转成**纯 VBA 代码**：用户把 `.bas` 导入
PowerPoint 运行 `BuildPoster`，得到**原生可编辑形状**（标题、面板、正文、
渐变背景、卡通形状），图标以**透明 PNG 图片**嵌入。与姊妹技能
sci-flowchart-vba（学术流程图）同源不同向——本技能面向**展板**，四项核心差异：

| 需求 | 本技能的做法 |
|---|---|
| 卡通图标 | ① 联网搜透明 PNG；② `scripts/icon_tool.py` 从原图抠图+清晰化；③ 引擎形状拼贴保底。`AddPicture` 嵌入，缺文件自动降级占位框 |
| 中文字体 | 双字体槽 `FONT_NAME`(西文) + `FONT_NAME_CN`(中文，写 NameFarEast)；`references/font-matching.md` 特征识别 + `scripts/font_check.py` 验证已安装；字号按像素测算，禁止一律 12pt |
| 数据图 | 先问用户偏好：空着 → 虚线占位框；简单画一个 → `AddBars` 原生形状迷你柱状图（无 Excel 依赖，两构建路径像素级一致） |
| 背景 | `AddBG` 纯色/上下渐变；装饰复杂 → `icon_tool.py` 抠干净区域贴回；背景烧死原文字 → 不透明面板盖板策略 |

## 两条路径

| 用户给了什么 | 模式 | 做法 |
|---|---|---|
| 一张展板图，要求"复刻/重绘/可编辑" | **模式① 复刻** | 程序化取色/字号 + 多模态识图 → 重建布局、素材、文字 |
| 参考展板图（风格）+ 内容清单/论文段落（内容） | **模式② 风格创作** | 从参考图取配色/字体/版式语言，内容换新 |

## 技能运行步骤（agent 照做）

1. **收输入**：展板图片（模式①）或参考图+内容（模式②）。问清数据图
   处理偏好（空着 / 简单画一个）。
2. **测量**（评审一的地基）：
   ```bash
   python scripts/measure_colors.py 原图.png --box 标题:x1,y1,x2,y2 --box 面板A:x1,y1,x2,y2
   python scripts/font_check.py 微软雅黑 思源黑体 华文行楷   # 按识别出的候选逐个验
   ```
   字号测算与字体特征识别见 `references/font-matching.md`。
3. **素材准备**（`references/icon-sourcing.md` 全流程）：
   - 图标：先联网搜透明 PNG（`icon_tool.py check` 验证）→ 搜不到合适的再
     从原图抠（`icon_tool.py cut --keyout auto --trim --scale 2 --sharpen 0.8`）
     → 都不行用引擎形状拼贴；
   - 背景：按 icon-sourcing.md 的背景策略表选 `AddBG` / 抠图贴回 / 盖板；
   - 产出 PNG 统一放 **输出目录 `assets\`**，`ASSET_DIR` 写它的绝对路径。
4. **给大模型下指令**：把原图 + 下方**「出码契约」**发给多模态模型，
   要求只回 `modPoster_Content.bas` 源码（可拆多模块）。
5. **拼装交付物**：`assets/modPoster_Engine.bas` **原样复制**进输出目录，
   写入模型产出的内容模块。
6. **生成 PPT**：
   ```bash
   python scripts/build_poster.py <输出目录> --out poster.pptx --export-png preview.png
   ```
   COM 优先（真 PowerPoint 绘制 + 导出 PNG 预览）；无 Office 自动回放
   （python-pptx 生成等价原生形状）。形状数 0 视为失败。
7. **两轮评审**（见下），通过后才准交付。
8. **交付**：`.pptx` 放第一位（`present_files`），附全部 `.bas`、`assets\`
   图标文件夹与导入说明。

---

## 出码契约（连同图片一起发给多模态模型）

你是一个"看展板写 PowerPoint VBA"的专家。给你一张展板图（模式②附带
风格参考图与内容清单）。直接产出 `modPoster_Content.bas` 源码，不要解释、
不要先转 SVG，只输出代码。

### 固定引擎 API（勿重写，全表见 references/poster-module-map.md）

- `AddNode sld, id, kind, x, y, w, h, fillC, lineC, lineW, dash, [text], [fontPt], [bold], [italic], [fontC], [cornerRatio], [fontName], [adj2]`
  万能形状节点；坐标为画布 px 左上角；`-1`=无填充/无边框；空槽位写 `""`/`-1`，不许 `,,`。
- `AddPicture sld, id, "file.png", x, y, w, h` — 图标图片（来自 ASSET_DIR，缺文件自动占位）。
- `AddBG sld, id, c1, c2` — 整版背景；`c2=-1` 纯色，否则上→下渐变；**最先调用**。
- `AddBars sld, id, x, y, w, h, title, "类别,…", "值,…", barC, lineC, txtC [, labelPt] [, titlePt]`
  — 迷你柱状图（原生形状）。
- `AddPath sld, id, "x1,y1;x2,y2", lineC, lineW, dash, arrowEnd` — 折线/箭头。
- `SetPara sld, id, "left", marginPx` 正文左对齐；`SetPartColor sld, id, nChars, colorC`
  前 n 字强调色；`BringToFront sld, id` 提到最前。
- kind 除流程图全家桶外还有卡通族：`star4/star5/star6/star8 heart cloud sun moon
  lightning smiley donut wave double_wave callout bubble`。

### 模块骨架（顺序固定）

```vba
Attribute VB_Name = "modPoster_Content"
Option Explicit
Option Base 0

' 1) 几何常量（11 个，缺一编译失败）
Public Const CANVAS_W_PX As Single = 1200
Public Const CANVAS_H_PX As Single = 675
Public Const PX_TO_IN As Single = 0.01007
Public Const OFFSET_X_IN As Single = 0.62
Public Const OFFSET_Y_IN As Single = 0.35
Public Const SLIDE_W_IN As Single = 13.333
Public Const SLIDE_H_IN As Single = 7.5
Public Const CUSTOM_SIZE As Boolean = False
Public Const FONT_NAME As String = "Microsoft YaHei"
Public Const FONT_NAME_CN As String = "微软雅黑"
Public Const ASSET_DIR As String = "C:\out\assets\"

' 2) 调色板（一律来自程序化取色，RGB 字面量）
Public Const BG_TOP As Long = RGB(255, 248, 235)
Public Const ACCENT As Long = RGB(64, 132, 220)

' 3) 唯一入口
Public Sub DrawAll(ByVal sld As Slide)
    DrawBackground sld          ' 背景必须最先
    DrawContent1 sld
End Sub
```

完整可运行示例：`references/example_content.bas`（渐变背景+面板+柱图+
图标+占位框+卡通装饰），拿不准时对齐它的写法。

### 几何常量怎么算

1. **选版面**：横版展板 → 16:9（13.333×7.5）；竖版展板/海报 → `7.5,10` 或
   按原图宽高比自定义（此时 `CUSTOM_SIZE=True`）。
2. `CANVAS_W_PX/CANVAS_H_PX` 取与原图同宽高比的整数画布（如 1200×675）。
3. `margin=0.35`；`availW=SLIDE_W-0.7`，`availH=SLIDE_H-0.7`；
   `PX_TO_IN=min(availW/CANVAS_W, availH/CANVAS_H)`；
   `OFFSET=margin+(avail-画布*PX_TO_IN)/2`。
4. 节点坐标在这张画布上按原图布局摆。

### 硬约束

- **文字保留源语言**：原图中文直接写中文；`.bas` 统一 UTF-8 + CRLF。
- **两行文字显式** `"A" & vbLf & "B"`，不靠自动折行。
- **坐标/颜色/文字来自识图与程序化测量**，禁止凭空猜；拿不准就近取实测值。
- **字号分层**：主标题 > 面板标题 > 正文 > 注释，同一层级统一，
  按字像素高换算（见 font-matching.md 的公式）。
- **中文字体写入 FONT_NAME_CN**，西文写 FONT_NAME；只用 font_check.py
  验证过的家族名，否则 PowerPoint 静默替换。
- **图标位置尽量贴合原图**；图标区域若原图是照片/复杂纹理，抠图失败的
  直接换搜索素材或形状拼贴，不要硬抠。
- **z-order = 调用顺序**：背景最先、图标画在所属面板之后（在其上层）。
- 节点+图片 > 18 或单模块 > 300 行 → 拆 `modPoster_Content2.bas`，
  `DrawAll` 依次调用。

---

## 最后一步：自动产出 PPT（不是只给 .bas）

`.bas` 只是中间产物，**最终交付物是 `.pptx`**。

```bash
python scripts/build_poster.py <输出目录> [--out poster.pptx] [--no-run]
                             [--replay] [--export-png preview.png]
```

- **COM 路径**（Windows + PowerPoint）：导入模块后真执行 `BuildPoster`，
  结果与用户按 F5 一致；`--export-png` 导出幻灯片 PNG 供评审。
- **回放路径**（无 Office）：解析 `AddBG/AddPicture/AddNode/AddPath/AddBars`
  调用，按**调用顺序**用 python-pptx 生成等价原生形状（含透明图片、渐变、
  东亚字体 a:ea）。缺失素材 → 与 COM 一致的虚线占位框。
- 输出目录最终必须同时有：`modPoster_Engine.bas`、`modPoster_Content*.bas`、
  **`*.pptx`**、`assets\` 图标、预览图。形状数为 0 = 失败，回查 `.bas`。
- `python-pptx`/`pywin32` 不在系统解释器时用隔离 venv：
  `C:/Users/Hello/.workbuddy/binaries/python/envs/default/Scripts/python.exe`
  （已预装 python-pptx、pywin32、Pillow、numpy）。

---

## 两轮评审（交付前必过）

### 评审一：与原图的程序化对照（.bas 常量 + 测量输出）

- **配色**：调色板常量必须能和 `measure_colors.py` 的输出对上号；
  强调色饱和度不许褪（原图鲜蓝就写实测 RGB，不许"差不多蓝"）；
  背景深浅、面板衬底同原图。
- **字体**：`font_check.py` 全部 `[可用]`；家族类别（黑/宋/楷/圆/行楷）
  与原图特征一致；中英分槽正确。
- **字号**：层级比例与原图一致，同一层级统一；用 compare_text.py
  带测量（源图传两遍）或 font-matching.md 公式核对。
- **素材**：图标 PNG 已验证透明（check 通过）且分辨率够（缩放后不糊）。

不通过 → 修常量/换素材 → 重新测量 → 通过才进评审二。

### 评审二：最终渲染视觉对照（preview.png / 导出 PNG vs 原图）

多模态逐项对照原图检查：

- [ ] 整体版式：区块位置、大小、留白与原图对应（允许 5% 内偏差）。
- [ ] 背景还原：纯色/渐变/装饰位置正确；盖板策略下原文字全部被盖住。
- [ ] 图标：位置、大小、清晰度合格；无占位框残留（除非素材确实缺失并已说明）。
- [ ] 文字：无溢出面板、无穿底、层级清楚；中文正确显示（无豆腐块）。
- [ ] 数据图：占位或柱图渲染正常，柱图数值/类目与内容一致。
- [ ] 卡通装饰（星星/爱心等）位置自然，不遮内容。

不通过 → 修 `.bas` → 重跑 `build_poster.py` → 再审。最多 3 轮，
仍不过就停下说明剩余问题，**不要带病交付**。

---

## 交付：PowerPoint 里怎么用

> 技能已通过 `build_poster.py` 直接给出 `.pptx`（形状已画好、图标已嵌入、
> 双击即编辑）。下面是"想从 `.bas` 重建"时的操作。

1. 打开 PowerPoint，`Alt+F11` 进 VBA 编辑器。
2. 每个 `.bas` → 右键工程 → Import File（引擎与内容都导）。
3. F5 运行 `BuildPoster`；想清空重画运行 `RemovePoster`。
4. 中文 Windows 手动导入前把 `.bas` 另存为 ANSI/GBK 可避免 VBE 乱码；
   用 `build_poster.py` 出图则无需此步。

## 自检清单（交付前过一遍）

- [ ] `modPoster_Engine.bas` 原样附带，未被改写。
- [ ] 内容模块声明全部 11 个常量（含 `FONT_NAME_CN`、`ASSET_DIR`）。
- [ ] `DrawAll` 存在且先调背景子过程；拆模块时全部 `DrawContentN` 被调用。
- [ ] 颜色全是 `RGB(r,g,b)`；无 `vbRed`；无空实参槽位（`,,`）。
- [ ] 全部 `.bas` 为 UTF-8 + CRLF。
- [ ] `font_check.py` 验证过全部字体名；中文字体写在 `FONT_NAME_CN`。
- [ ] 图标 PNG 经 `icon_tool.py check` 验证；`ASSET_DIR` 为绝对路径且以 `\` 结尾。
- [ ] **`build_poster.py` 已跑通，`.pptx` 里有形状**（非空板）。
- [ ] 评审一、评审二均已通过（保留 preview.png 供用户对照）。
- [ ] 交付物：`.pptx`（第一位）+ 全部 `.bas` + `assets\` 图标 + 导入说明。

## 常见问题排查

| 现象 | 原因 | 处理 |
|---|---|---|
| 编译报变量未定义（如 `ASSET_DIR`） | 内容模块漏常量 | 补齐 11 个 `Public Const` |
| 中文显示为宋体/豆腐块 | `FONT_NAME_CN` 没写或字体未安装 | font_check.py 验证后写入；勿用未装字体 |
| 图标位置是虚线灰框 | 素材文件没找到 | 核对 `ASSET_DIR` 绝对路径与文件名大小写；重跑 |
| 抠图后边缘白边/黑边 | tol 过大或底色渐变 | 降 tol、bbox 贴内侧、`--trim`；顽固白边换路线 A 素材 |
| 放大后图标发糊 | 原图分辨率不足 | scale≤3 + sharpen；仍糊换素材或形状拼贴 |
| 柱状图与预览不一致 | 单侧改过 AddBars 几何 | 引擎与回放路径几何必须同步改（见 module-map） |
| `.pptx` 空白 | COM 异常被吞 / 漏 `DrawAll` | 看脚本日志；`--replay` 兜底并回查 `.bas` |
| 重建报 `PermissionError` | pptx 被预览占用 | `--out poster_v2.pptx` 换名输出 |
| 回放路径渐变方向不对 | 极少数 python-pptx 版本差异 | 评审二会发现；必要时改纯色或手动调 gradient_angle |

## 资源

- `assets/modPoster_Engine.bas` — 固定引擎（映射/形状/图片/背景/柱图/字体/幂等/入口）
- `references/poster-module-map.md` — API 全表、kind→mso 编号、双路径一致性约定
- `references/example_content.bas` — 可运行示例（覆盖全部新能力）
- `references/font-matching.md` — 中文字体特征识别 + 字号测算公式 + 回退策略
- `references/icon-sourcing.md` — 图标三路线（搜索/抠图/形状拼贴）+ 背景策略表
- `scripts/build_poster.py` — `.bas` → 真实 `.pptx`（COM 优先，回放兜底，--export-png）
- `scripts/icon_tool.py` — probe / check / cut（抠底、去白边、放大、锐化）
- `scripts/font_check.py` — 已安装字体枚举与匹配回退
- `scripts/measure_colors.py` — 通用取色（主色/区域底色/文字色/强调色）
- `scripts/compare_text.py` — 文字带字号校准（评审一）
