# modPoster_Engine API 地图与内容模块契约

引擎 `assets/modPoster_Engine.bas` 由 skill 固定提供，**内容模块只调用、不重写**。
COM 路径（PowerPoint VBA）与回放路径（`build_poster.py` + python-pptx）解析同一套
调用，几何语义一致。

## 内容模块必须声明的全局常量（缺一编译失败）

| 常量 | 类型 | 说明 |
|---|---|---|
| `CANVAS_W_PX` / `CANVAS_H_PX` | Single | 布局画布宽/高（px，与原图同宽高比） |
| `PX_TO_IN` | Single | px→英寸比例 `min(availW/CANVAS_W, availH/CANVAS_H)` |
| `OFFSET_X_IN` / `OFFSET_Y_IN` | Single | 画布在幻灯片上的左上偏移（居中公式） |
| `SLIDE_W_IN` / `SLIDE_H_IN` | Single | 版面尺寸（宽图 16:9=13.333×7.5） |
| `CUSTOM_SIZE` | Boolean | 非 16:9 版面时 True |
| `FONT_NAME` | String | 西文字体家族名 |
| `FONT_NAME_CN` | String | 中文字体家族名（写入 NameFarEast 槽） |
| `ASSET_DIR` | String | 图标素材文件夹**绝对路径**，以 `\` 结尾 |

必须定义：`Public Sub DrawAll(ByVal sld As Slide)`，内部依次调用全部
`DrawContentN` / `DrawBackground`。**背景子过程必须最先调用**（z-order = 调用顺序）。

## API 全表

### AddNode —— 万能形状节点
```vba
AddNode sld, id, kind, x, y, w, h, fillC, lineC, lineW, dash, _
        [text], [fontPt], [bold], [italic], [fontC], [cornerRatio], [fontName], [adj2]
```
- 坐标为画布 px，(x,y) 左上角；颜色 `RGB(r,g,b)`，传 `-1` 表示无填充/无边框。
- `fontName` 省略时西文用 `FONT_NAME`、中文用 `FONT_NAME_CN`；显式给出则两者都用它。
- `cornerRatio`：圆角矩形=圆角比例(0~0.5)；`stadium` 自动 0.5；star 族=内半径比例。
- `adj2`：块箭头的箭头长度比例。
- 无实参槽位必须写值（`""` / `-1` / `0`），不能留空产生 `,,`。

### AddPicture —— 图标/装饰图片层
```vba
AddPicture sld, id, "battery.png", x, y, w, h
```
- 文件解析顺序：原样路径 → `ASSET_DIR & 文件名`。推荐 PNG（支持透明）。
- w/h 为画布 px；两者都 ≤0 时按图片原始尺寸（两路径 DPI 语义略有差异，尽量显式给）。
- **文件缺失不报错**：两条路径都降级为灰色虚线占位框（内标 `IMG 文件名`）。
- 图片是位图：颜色不能像形状一样改，尺寸可缩放。

### AddBG —— 整版背景（必须最先调用）
```vba
AddBG sld, "bg", BG_TOP, BG_BOT    ' 上→下双色渐变
AddBG sld, "bg", BG_TOP, -1        ' 纯色
```

### AddBars —— 原生形状拼的迷你柱状图（无 Excel 依赖）
```vba
AddBars sld, id, x, y, w, h, title, "类别1,类别2", "v1,v2", barC, lineC, txtC _
        [, labelPt] [, titlePt] [, fontName]
```
内部几何（画布 px，回放路径逐字节复刻，勿单侧改动）：
- 标题带 `y..y+0.16h`（title 为空则无）；类目带 `y+0.86h..y+h`；基线 `y+0.86h`。
- `slot=w/n`，柱宽 `0.6*slot`，柱 i 左缘 `x+slot*i+0.2*slot`。
- 柱高 `v/vmax*plotH*0.90`；数值标签在柱上方 `0.08h` 带内；基线为一条线段。
- 生成的子形状 id：`id_t / id_bN / id_vN / id_cN / id_axis`（可被 SetPara 引用）。


### AddFormula —— 原生 Office 公式（LaTeX）

AddFormula sld, "fx", x, y, w, h, "CESI = \sqrt[3]{EHI \times (1-ERI) \times ESI}", _
    "CESI = (EHI x (1-ERI) x ESI)^(1/3)", 13, INK

- 建一个透明文本框占位，`build_poster.py` 在保存后把它替换为**原生 OMML 公式**
  （COM 路径同样后处理；VBA 手动 F5 只能看到线性 fallback 文本）。
- `tex` 用双引号包裹、**单反斜杠**（VBA 不转义反斜杠；单引号是行注释）。
- `fallBack` 是给 VBE/无 math 依赖环境的线性版，不要事后手改（注入按它找形状）。
- 字号：公式宽 ≈ fontPt × 1.83px/字符 × 字符数；溢出就缩字号或加宽 w。
- 预览：matplotlib mathtext（STIX）渲染同一 LaTeX，失败降级线性文本。
- 依赖：`pip install latex2mathml mathml2omml`。

### AddPath —— 折线/箭头
```vba
AddPath sld, id, "x1,y1;x2,y2;x3,y3", lineC, lineW, dash, arrowEnd
```

### 文字后处理
```vba
SetPara sld, id, "left"|"center"|"right", marginPx   ' 正文左对齐+左边距
SetPartColor sld, id, nChars, RGBColor               ' 前 nChars 字变强调色+加粗
BringToFront sld, id                                  ' 提到最前（回放路径为 no-op）
ShapeById(sld, id) As Shape                           ' 引擎内供取形状
```

### 入口
`BuildPoster`（画图）/ `RemovePoster`（清空本 TAG 全部形状）。

## 形状 kind → mso 编号（已核对 MSO_AUTO_SHAPE_TYPE）

| kind | mso | | kind | mso |
|---|---|---|---|---|
| rect | 1 | | star4 | 91 |
| round_rect / stadium | 5 | | star5 | 92 |
| parallelogram | 2 | | star8 | 93 |
| trapezoid | 3 | | star6 | 147 |
| diamond | 4 | | heart | 21 |
| oval | 9 | | cloud | 179 |
| hexagon | 10 | | sun | 23 |
| can / cylinder | 13 | | moon | 24 |
| smiley | 17 | | lightning | 22 |
| donut | 18 | | wave | 103 |
| pentagon(箭头形) | 51 | | double_wave | 104 |
| pentagon_reg | 12 | | callout(圆角气泡) | 106 |
| chevron | 52 | | bubble(椭圆气泡) | 107 |
| document | 61 | | arc | 25 |
| right/left/up/down_arrow | 33/34/35/36 | | | |

## 两条路径的一致性约定

- **绘制顺序 = 调用顺序**（回放路径保持 ops 顺序，z-order 与 VBA 一致）。
- 回放路径 `BringToFront` 为 no-op（顺序已定）；确需"图标压在后面面板上"，
  在 VBA 里也可通过调整 DrawAll 里的调用顺序实现。
- 缺失素材 → 两路径一致的虚线占位框。
- `.bas` 一律 **UTF-8 + CRLF**；中文直接写，不翻译。

## 拆模块规则

节点+图片 > 18 个、或单模块 > 300 行 → 拆 `modPoster_Content2.bas`…，
`DrawAll` 依次调用全部 `DrawContentN`；背景逻辑放 `DrawBackground` 并保持最先调用。
