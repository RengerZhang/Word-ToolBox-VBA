Attribute VB_Name = "编号_1_导入各级大纲样式"
Option Explicit

'（一）表格样式名常量（模块级，供全模块任何过程复用）
Private Const STYLE_TABLE_NORMAL As String = "标准表格样式"   ' 普通表：仅作为标记，不改外观
Private Const STYLE_TABLE_PIC    As String = "图片定位表"     ' 图片表：无框线 + 0 内边距

Sub 一键设置全部样式()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Call 一键设置正文格式
    Call 设置标题样式(doc, "标题 1", 18, wdOutlineLevel1, 0.5, 0.5, wdAlignParagraphCenter)
    Call 设置标题样式(doc, "标题 2", 14, wdOutlineLevel2, 0.5, 0, wdAlignParagraphLeft)
    Call 设置标题样式(doc, "标题 3", 12, wdOutlineLevel3, 0, 0, wdAlignParagraphLeft)
    Call 设置标题样式(doc, "标题 4", 12, wdOutlineLevel4, 0, 0, wdAlignParagraphLeft)
    Call 创建条款项样式(doc, "条样式【1）】", wdOutlineLevelBodyText)
    Call 创建条款项样式(doc, "款样式【（1）】", wdOutlineLevelBodyText)
    Call 创建条款项样式(doc, "项样式【①】", wdOutlineLevelBodyText)
    Call EnsureStandardTableStyle
    Call 导入图片标题_子图样式
    
    '========================
    ' 4. 表格标题 / 图片标题
    '========================
    Dim styleTableCaption As Style, stylePicCaption As Style, stylePicPara As Style   ' ← 新增 stylePicPara
    
    On Error Resume Next
    Set styleTableCaption = doc.Styles("表格标题")
    If styleTableCaption Is Nothing Then
        Set styleTableCaption = doc.Styles.Add(name:="表格标题", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    With styleTableCaption.Font
        .NameFarEast = "黑体"
        .NameAscii = "Times New Roman"
        .bold = True
        .Size = 10.5
    End With
    With styleTableCaption.ParagraphFormat
        .outlineLevel = wdOutlineLevelBodyText
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpace1pt5   '1.5倍行距
        .SpaceBefore = Application.LinesToPoints(0)   ' 段前 0 行
        .SpaceAfter = Application.LinesToPoints(0)    ' 段后 0 行
        ' ===（新增）与下段同页===
        .KeepWithNext = True
    End With
    
    ' 图片标题 = 继承表格标题
    On Error Resume Next
    Set stylePicCaption = doc.Styles("图片标题")
    If stylePicCaption Is Nothing Then
        Set stylePicCaption = doc.Styles.Add(name:="图片标题", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    stylePicCaption.BaseStyle = "表格标题"
    With stylePicCaption.ParagraphFormat
        .KeepWithNext = False
        .outlineLevel = wdOutlineLevelBodyText
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpace1pt5   '1.5倍行距
        .SpaceBefore = Application.LinesToPoints(0)   ' 段前 0 行
        .SpaceAfter = Application.LinesToPoints(0)    ' 段后 0 行
    End With
    
    '========================
    ' 4+. 图片段落样式：图片格式（独立样式）
    '========================
    On Error Resume Next
    Set stylePicPara = doc.Styles("图片格式")
    If stylePicPara Is Nothing Then
        Set stylePicPara = doc.Styles.Add(name:="图片格式", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    
    ' 独立样式：不继承任何样式
    stylePicPara.BaseStyle = ""
    stylePicPara.NextParagraphStyle = doc.Styles("正文")
    
    ' 仅设置你要求的属性；其余保持默认
    With stylePicPara.ParagraphFormat
        .outlineLevel = wdOutlineLevelBodyText     ' 正文级
        .LeftIndent = 0                            ' 无缩进
        .RightIndent = 0
        .FirstLineIndent = 0                       ' 无首行缩进
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphCenter        ' 居中
        .KeepWithNext = True                       ' 与下段同页
        ' 不设置行距/段前段后/边框底纹/制表位等，保持默认
    End With
    
    '========================
    ' 5. 表格样式
    '========================
    Dim tblStyleStd As Style, tblStylePic As Style
    
    ' 图片定位表
    On Error Resume Next
    Set tblStylePic = doc.Styles("图片定位表")
    If tblStylePic Is Nothing Then
        Set tblStylePic = doc.Styles.Add(name:="图片定位表", Type:=wdStyleTypeTable)
    End If
    On Error GoTo 0
    With tblStylePic.Table
        .Borders.enable = False
        .alignment = wdAlignRowCenter
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
    End With
    With tblStylePic.ParagraphFormat
        .KeepWithNext = True
        .FirstLineIndent = 0
        .LeftIndent = 0
        .RightIndent = 0
    End With
    
    
    '========================
    ' TOC 1~3：首行缩进关掉 + 单倍行距（逐个设置，便于以后独立调整）
    '========================
    
    '（一）TOC 1
    With doc.Styles(wdStyleTOC1).ParagraphFormat
        .FirstLineIndent = 0                 ' ① 首行缩进（磅）清零
        .CharacterUnitFirstLineIndent = 0    ' ② 首行缩进（字符）清零
        .LineSpacingRule = wdLineSpaceSingle ' ③ 行距设为“单倍”
    End With
    
    '（二）TOC 2
    With doc.Styles(wdStyleTOC2).ParagraphFormat
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
    
    '（三）TOC 3
    With doc.Styles(wdStyleTOC3).ParagraphFormat
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
    
    '——TOC 标题1：首行缩进关掉 + 单倍行距 + 宋体/加粗/小四
    With ActiveDocument.Styles("TOC 标题")
        '（一）段落：关首行缩进，行距=单倍
        With .ParagraphFormat
            .FirstLineIndent = 0                   ' ① 首行缩进（磅）清零
            .CharacterUnitFirstLineIndent = 0      ' ② 首行缩进（字符）清零
            .LineSpacingRule = wdLineSpaceSingle   ' ③ 行距设为“单倍”
            .alignment = wdAlignParagraphCenter
        End With
        '（二）字体：宋体 + 加粗 + 小四（12pt）
        With .Font
            .NameFarEast = "宋体"                  ' ④ 中文字体=宋体
            .Size = 18                             ' ⑤ 小四号=12 磅
            .bold = True                           ' ⑥ 加粗
            .Color = wdColorBlack
            .NameAscii = "Times New Roman"
        End With
    End With
    
    
    
    MsgBox "所有样式一键导入完成！"
End Sub
'========================
' 辅助函数：设置标题样式（可传段前/段后“行”、对齐方式）
' 参数：
'   styleName   标题样式名（如“标题 1”）
'   fontSize    字号（pt）
'   olLevel     大纲级别
'   beforeLines 段前(行)  —— 可选，默认0
'   afterLines  段后(行)  —— 可选，默认0
'   align       对齐方式  —— 可选，默认左对齐
'========================
Private Sub 设置标题样式( _
    doc As Document, ByVal styleName As String, ByVal fontSize As Single, _
    ByVal olLevel As WdOutlineLevel, _
    Optional ByVal beforeLines As Single = 0, _
    Optional ByVal afterLines As Single = 0, _
    Optional ByVal align As WdParagraphAlignment = wdAlignParagraphLeft)

    Dim st As Style
    Set st = doc.Styles(styleName)   ' 假定内置标题样式已存在

    '（一）字体
    With st.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .bold = True
        .Size = fontSize
    End With

    '（二）段落
    With st.ParagraphFormat
        .outlineLevel = olLevel
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = align
        ' 将“行”换算为磅
'        .SpaceBefore = Application.LinesToPoints(beforeLines)
'        .SpaceAfter = Application.LinesToPoints(afterLines)
        .SpaceBefore = beforeLines
        .SpaceAfter = afterLines
    End With

    '（三）下一段样式回到正文
    st.NextParagraphStyle = doc.Styles("正文")
End Sub


'========================
' 辅助函数：创建条款项样式
'========================
Private Sub 创建条款项样式(doc As Document, ByVal styleName As String, ByVal olLevel As WdOutlineLevel)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    If st Is Nothing Then
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    
    st.BaseStyle = "正文"
    With st.ParagraphFormat
        .outlineLevel = olLevel
        .CharacterUnitFirstLineIndent = 2
    End With
    
    st.NextParagraphStyle = ActiveDocument.Styles("正文")
End Sub
Private Sub 一键设置正文格式()
    Selection.Style = ActiveDocument.Styles("正文")
    With ActiveDocument.Styles("正文").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 12
        .bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    
    With ActiveDocument.Styles("正文").ParagraphFormat
        .CharacterUnitFirstLineIndent = 2
        .outlineLevel = wdOutlineLevelBodyText
        .LeftIndent = CentimetersToPoints(0) ' 左侧0字符
        .RightIndent = CentimetersToPoints(0) ' 右侧0字符
        .SpaceBefore = 0 ' 段前0磅
        .SpaceAfter = 0 ' 段后0磅
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
    End With
    
End Sub


Public Sub 导入两种表格样式_仅表格层()
    Dim doc As Document: Set doc = ActiveDocument
    Dim styStd As Style, styPic As Style

    '（一）确保存在“标准表格样式”（普通表的标记用；不做任何视觉设置）
    Set styStd = EnsureTableStyleOnly(doc, STYLE_TABLE_NORMAL)
    ' ——不设置 .Table/.ParagraphFormat/.Font，完全保持默认，由后续格式化流程统一处理——

    '（二）确保存在“图片定位表”（图片表专用）
    Set styPic = EnsureTableStyleOnly(doc, STYLE_TABLE_PIC)
    With styPic.Table
        '（1）关闭全部框线
        .Borders.enable = False
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone

        '（2）行高保持“自适应”（默认即为自动行高，这里不显式设置）

        '（3）单元格四个方向的距离均为 0
        .TopPadding = 0
        .BottomPadding = 0
        .LeftPadding = 0
        .RightPadding = 0

        '（4）不设置段落/字符相关属性（水平/垂直居中等留待后续流程处理）
    End With
End Sub
' === 标准化表格样式：若存在则修改；不存在则创建（段落样式）===
Public Sub EnsureStandardTableStyle()
    Dim stdStyle As Style

    ' 1) 尝试获取已存在的样式
    On Error Resume Next
    Set stdStyle = ActiveDocument.Styles("标准化表格样式")
    On Error GoTo 0

    ' 2) 不存在则创建；存在但类型非“段落样式”则重建为段落样式
    If stdStyle Is Nothing Then
        Set stdStyle = ActiveDocument.Styles.Add( _
            name:="标准化表格样式", Type:=wdStyleTypeParagraph)
    ElseIf stdStyle.Type <> wdStyleTypeParagraph Then
        stdStyle.Delete
        Set stdStyle = ActiveDocument.Styles.Add( _
            name:="标准化表格样式", Type:=wdStyleTypeParagraph)
    End If

    ' 3) 设置样式属性
    With stdStyle
        .AutomaticallyUpdate = False

        ' 尝试设为“独立样式”；若不被接受则回退到 Normal
        On Error Resume Next
        .BaseStyle = ""                                ' 不继承正文
        If Err.Number <> 0 Then
            Err.Clear
            .BaseStyle = ActiveDocument.Styles(wdStyleNormal)
        End If
        On Error GoTo 0

        ' ====== 字体 ======
        With .Font
            .NameFarEast = "宋体"                      ' 中文
            .NameAscii = "Times New Roman"            ' 英文/数字
            .Size = 10.5                               ' 五号
            .bold = False
            .Color = wdColorAutomatic
        End With

        ' ====== 段落 ======
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceSingle       ' 单倍行距
            .SpaceBefore = 0                           ' 段前 0
            .SpaceAfter = 0                            ' 段后 0
            .LeftIndent = 0                            ' 左缩进
            .RightIndent = 0                           ' 右缩进
            .FirstLineIndent = 0                       ' 首行缩进
            .CharacterUnitFirstLineIndent = 0
            .alignment = wdAlignParagraphCenter        ' 居中
            '.NoSpaceBetweenParagraphsOfSameStyle = True ' 可按需启用
        End With
    End With
End Sub
'========================================================
' 工具：仅保证“表格样式”的存在与类型正确（不改任何段落/字符属性）
' - 若同名样式存在但类型不是“表格样式”，则删除后以“表格样式”重建
'========================================================
Private Function EnsureTableStyleOnly(ByVal doc As Document, ByVal styleName As String) As Style
    Dim st As Style

    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0

    If Not st Is Nothing Then
        If st.Type <> wdStyleTypeTable Then
            '（一）同名但类型不对，删除并重建为“表格样式”
            st.Delete
            Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
        End If
    Else
        '（二）不存在则新建
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
    End If

    Set EnsureTableStyleOnly = st
End Function
'==============================
'【新增】图片子图名样式（独立于“图片标题”）
'  默认：宋体、12pt、非加粗、单倍行距、居中、无继承、关对齐到网格
'==============================
Public Sub 导入图片标题_子图样式()
    '（一）创建/更新样式本体
    Call EnsureParagraphStyle( _
        styleName:="图片标题-子图", _
        nameCN:="宋体", nameEN:="Times New Roman", _
        ptSize:=10.5, isBold:=False, _
        lineRule:=wdLineSpace1pt5, align:=wdAlignParagraphCenter)

    '（二）补充通用属性：无继承、关闭“对齐到网格”、段前后/缩进清零
    With ActiveDocument.Styles("图片标题-子图")
        .AutomaticallyUpdate = False
        On Error Resume Next
        .BaseStyle = "图片标题"                           ' 不基于任何样式（无继承）
        On Error GoTo 0
        .Font.bold = False
        .Font.Size = 10.5
        
        With .ParagraphFormat
            .SpaceBefore = 0: .SpaceAfter = 0
            .LeftIndent = 0: .RightIndent = 0
            .FirstLineIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .DisableLineHeightGrid = False
        End With
    End With
End Sub

'===============================
'  （一）先安全获取样式对象（捕获获取失败）
'  （二）分两段判断：若 Nothing → 新建；否则再判类型
'  （三）最后再设置字体/段落属性
'===============================
Public Sub EnsureParagraphStyle( _
    ByVal styleName As String, _
    ByVal nameCN As String, ByVal nameEN As String, _
    ByVal ptSize As Single, ByVal isBold As Boolean, _
    ByVal lineRule As WdLineSpacing, ByVal align As WdParagraphAlignment)

    Dim st As Style

    '（一）安全获取样式对象
    On Error Resume Next
    Set st = ActiveDocument.Styles(styleName)
    If Err.Number <> 0 Then
        Err.Clear
        Set st = Nothing
    End If
    On Error GoTo 0

    '（二）不存在 or 类型不对 → 新建为“段落样式”
    If st Is Nothing Then
        Set st = ActiveDocument.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    Else
        ' 注意：VBA 无短路求值，必须分支后再访问 st.Type
        If st.Type <> wdStyleTypeParagraph Then
            On Error Resume Next
            st.Delete                      ' 同名但类型不对（如表格/字符样式），删掉重建
            On Error GoTo 0
            Set st = ActiveDocument.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
        End If
    End If

    '（三）统一设置样式属性
    With st
        .AutomaticallyUpdate = False

        ' 基础样式尽量指向“正文/常规”，不同环境容错
        On Error Resume Next
        .BaseStyle = ""                 ' ← 关键：不基于任何样式
        Err.Clear
        On Error GoTo 0

        ' 字体（中英文分别设置）
        With .Font
            .NameFarEast = nameCN
            .NameAscii = nameEN
            .Size = ptSize
            .bold = isBold
            .Color = wdColorAutomatic
        End With

        ' 段落：对齐 + 1.5倍行距（用枚举 wdLineSpace1pt5）
        With .ParagraphFormat
            .alignment = align
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = lineRule
            .FirstLineIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .DisableLineHeightGrid = True
        End With
    End With
End Sub

