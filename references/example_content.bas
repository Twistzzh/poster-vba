Attribute VB_Name = "modPoster_Content"
Option Explicit
Option Base 0

' ============================================================================
' poster-vba 示例内容模块 —— 一个迷你比赛展板，覆盖全部新增能力：
' 渐变背景 / 图片图标 / 圆角面板 / 简易柱状图 / 虚线数据图占位 /
' 卡通形状(星星、爱心) / 左对齐正文。真实任务里由多模态模型按原图生成。
' ============================================================================

' 1) 几何常量（引擎引用，缺一会编译失败）
Public Const CANVAS_W_PX As Single = 1200
Public Const CANVAS_H_PX As Single = 675
Public Const PX_TO_IN As Single = 0.01007
Public Const OFFSET_X_IN As Single = 0.62
Public Const OFFSET_Y_IN As Single = 0.35
Public Const SLIDE_W_IN As Single = 13.333
Public Const SLIDE_H_IN As Single = 7.5
Public Const CUSTOM_SIZE As Boolean = False
Public Const FONT_NAME As String = "Microsoft YaHei"   ' 西文字体
Public Const FONT_NAME_CN As String = "微软雅黑"        ' 中文字体（东亚字体槽）
' 图标所在文件夹：必须写成输出目录下的绝对路径，以 \ 结尾
Public Const ASSET_DIR As String = "C:\out\assets\"

' 2) 调色板（从原图程序化取色，一律 RGB 字面量）
Public Const BG_TOP As Long = RGB(255, 248, 235)
Public Const BG_BOT As Long = RGB(255, 236, 210)
Public Const ACCENT As Long = RGB(64, 132, 220)
Public Const ACCENT2 As Long = RGB(255, 158, 66)
Public Const INK As Long = RGB(51, 51, 51)
Public Const ALT As Long = RGB(224, 239, 252)
Public Const ALT2 As Long = RGB(255, 238, 220)
Public Const ALT3 As Long = RGB(232, 245, 233)
Public Const LINE As Long = RGB(150, 150, 150)
Public Const WHITE As Long = RGB(255, 255, 255)

' 3) 唯一入口（BuildPoster 会调用它）
Public Sub DrawAll(ByVal sld As Slide)
    DrawBackground sld
    DrawContent1 sld
    DrawContent2 sld
End Sub

' 4) 背景（必须是 DrawAll 里最先调用的子过程）
Public Sub DrawBackground(ByVal sld As Slide)
    AddBG sld, "bg", BG_TOP, BG_BOT          ' c2 >= 0 -> 上下渐变
    ' 纯色背景写法: AddBG sld, "bg", BG_TOP, -1
End Sub

Public Sub DrawContent1(ByVal sld As Slide)
    ' 标题横幅
    AddNode sld, "title", "round_rect", 30, 25, 1140, 80, ACCENT, -1, 1, LINE_SOLID, _
                "2026 城市环保创意大赛 展板", 26, True, False, WHITE, 0.18

    ' 面板 A：研究背景
    AddNode sld, "panelA", "round_rect", 30, 130, 540, 250, ALT, -1, 1, LINE_SOLID, "", 11, False, False, INK, 0.08
    AddNode sld, "headA", "rect", 55, 145, 490, 40, -1, -1, 1, LINE_SOLID, "研究背景", 16, True, False, RGB(31, 78, 150)
    AddNode sld, "bodyA", "rect", 55, 195, 490, 165, -1, -1, 1, LINE_SOLID, _
                "城市垃圾分类推行三年来，居民参与率提升明显，" & vbLf & _
                "但投放准确率仍不足六成。" & vbLf & _
                "本作品设计一套智能督导装置，用图像识别" & vbLf & _
                "实时纠正错误投放。", 12, False, False, INK
    SetPara sld, "bodyA", "left", 8

    ' 面板 B：数据一览（简易柱状图）
    AddNode sld, "panelB", "round_rect", 630, 130, 540, 250, ALT2, -1, 1, LINE_SOLID, "", 11, False, False, INK, 0.08
    AddNode sld, "headB", "rect", 655, 145, 490, 40, -1, -1, 1, LINE_SOLID, "试点数据", 16, True, False, RGB(200, 100, 20)
    AddBars sld, "bars", 660, 195, 480, 170, "", "调研,设计,制作,答辩", "3.2,5.1,4.4,2.8", _
                ACCENT, LINE, INK, 10

    ' 面板 C：成果亮点
    AddNode sld, "panelC", "round_rect", 30, 405, 760, 240, ALT3, -1, 1, LINE_SOLID, "", 11, False, False, INK, 0.08
    AddNode sld, "headC", "rect", 55, 420, 710, 40, -1, -1, 1, LINE_SOLID, "成果亮点", 16, True, False, RGB(56, 128, 82)
    AddNode sld, "bodyC", "rect", 55, 470, 710, 155, -1, -1, 1, LINE_SOLID, _
                "识别准确率 94.2%，督导响应时间 0.8 秒。" & vbLf & _
                "已在 2 个社区完成 6 周试点，覆盖 412 户居民。", 12, False, False, INK
    SetPara sld, "bodyC", "left", 8
End Sub

Public Sub DrawContent2(ByVal sld As Slide)
    ' 卡通电池图标（找不到文件时自动降级为虚线占位框）
    AddPicture sld, "battery", "battery.png", 830, 430, 130, 180

    ' 数据图占位（数据图可以先空着）
    AddNode sld, "chartph", "rect", 990, 430, 180, 150, WHITE, LINE, 1, LINE_DASH, _
                "数据图" & vbLf & "占位", 11, False, False, RGB(120, 120, 120)

    ' 装饰：五角星 + 爱心（展板的活泼元素用形状表达，避开文字/标题）
    AddNode sld, "star1", "star5", 1095, 600, 30, 30, RGB(255, 205, 66), -1, 1, LINE_SOLID, "", 8
    AddNode sld, "star2", "star5", 1055, 622, 22, 22, RGB(255, 205, 66), -1, 1, LINE_SOLID, "", 8
    AddNode sld, "heart1", "heart", 1130, 610, 42, 40, RGB(240, 98, 146), -1, 1, LINE_SOLID, "", 8

    ' 面板 A -> 面板 B 的流程箭头
    AddPath sld, "arrowAB", "570,255;630,255", ACCENT, 2, LINE_SOLID, True
End Sub
