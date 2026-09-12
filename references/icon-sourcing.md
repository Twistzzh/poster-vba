# 卡通图标与背景素材获取指南

展板的"活泼感"主要来自卡通图标（电池、灯泡、叶子、云朵…）。三条素材路线，
按顺序尝试，全部失败再降级为形状拼贴。

## 路线 A：联网搜现成透明 PNG（首选，最快）

1. **搜索词模板**（中英都试）：
   - `电池 卡通 插画 免抠 png` / `battery cartoon icon transparent png`
   - `灯泡 创意 扁平 透明背景` / `环保 叶子 图标 透明`
2. **优先站点**：iconfont.cn（注意看授权标注）、flaticon、icons8、
   pngwing/cleanpng 等免抠站；比赛公开使用优先选**免费可商用**素材。
3. **下载**：`curl -L -o battery.png "<直链>"`。直链拿不到时抓页面里的
   CDN 图址；带防盗链的站补 `-H "Referer: <页面地址>"`。
4. **下载后必须验证**（很多"免抠"图其实带白底）：
   ```bash
   python scripts/icon_tool.py check battery.png
   ```
   - 结论"已是透明背景" → 直接用（必要时 `--trim` 裁掉透明边）。
   - 结论"需要 keyout" → 走路线 B 的 cut 流程去底。

## 路线 B：从原图上抠下来 + 清晰化（原图里就有该图标时）

1. 先 `probe` 摸底：
   ```bash
   python scripts/icon_tool.py probe 原图.png
   ```
   看尺寸、边界中位色（决定 tol 起点）、是否自带 alpha。
2. 用看图坐标估一个**贴着图标内侧**的 bbox，宁小勿大（框进周围文字/装饰
   会让洪泛抠底失败）：
   ```bash
   python scripts/icon_tool.py cut 原图.png battery.png --bbox 830,430,960,610 --keyout auto --tol 40 --trim --scale 2 --sharpen 0.8
   ```
3. 逐参数微调：
   - 底色有渐变/阴影 → `--tol` 加大到 60~80；
   - 图标内部有同底色孔洞被误删 → 改 `--keyout R,G,B`（精确底色，全图匹配）；
   - 抠掉比例 >90% 会有警告 → 多半是 bbox 框错或 tol 过大；
   - 输出偏糊 → `--scale 3` + `--sharpen 1.0`；**超过 3 倍放大不会更清晰**，
     小图标糊得救不回来就换路线 A 或 C。
4. 产出直接被 `AddPicture` 引用；把文件放进输出目录 `assets\`，
   并让内容模块的 `ASSET_DIR` 指向它的绝对路径。

## 路线 C：形状拼贴（保底，风格最统一）

简单卡通物可以全部用引擎形状拼出来，纯矢量、可编辑、永不吃像素：
- **电池** = 竖 `round_rect`(壳, 深色描边) + 顶部小 `rect`(电极) + 内嵌浅绿
  `rect`(电量) + 可选 `lightning`(闪电标识)；
- **奖杯** = `trapezoid`(杯身) + 两个 `arc`(耳) + `rect`(底座)；
- **云朵** = 引擎自带 `cloud`；星星 `star5`、爱心 `heart`、太阳 `sun`、
  气泡 `bubble`/`callout` 都有现成 kind（见 module-map 的 kind 全表）。

## 展板背景策略（对应"背景从原图扣或保持基本配色"）

| 原图背景形态 | 做法 |
|---|---|
| 纯色 | `AddBG sld,"bg",颜色,-1` |
| 上下/左右渐变 | `AddBG` 双色渐变（取色用 `measure_colors.py` 的 bg_guess + 对侧中位色） |
| 角落装饰（云、波浪、彩带），中部干净 | 整块或角部 `icon_tool.py cut`（`--keyout none` 只裁剪）成 PNG，`AddPicture` 贴回四角 |
| 背景大面积装饰图案 | 优先抠**无文字**的干净区域；实在无法分离 → 整幅原图做背景 `AddPicture`(0,0,满画布) + 所有面板用不透明填充盖住烧在背景里的原文字 |
| 背景里烧死了原文字 | 上面那种"盖板"策略；面板颜色从原图取，保证观感一致 |

> 注意：整幅原图当背景会把原文字带进 PPT（不可编辑）。只有在背景装饰
> 复杂、且面板能盖住全部文字区时才用，交付说明里要写明。
