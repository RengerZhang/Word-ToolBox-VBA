Attribute VB_Name = "样式_四级标题_表格标题"
Private Sub 获取当前多级列表模板()
    Dim oDoc As Document
    Set oDoc = Word.ActiveDocument
    Dim oRng As Range
    Dim oList As List
    Dim oListFormat As ListFormat
    Dim oP As Paragraph
    Set oRng = Word.Selection.Range
    With oRng
        '获取当前选中内容所在的第一个列表项目编号的字符串,比如"2.3.1"
       MsgBox .ListFormat.ListType
    End With
End Sub

Sub 消除前四级标题缩进量()
    Dim 文档 As Document
    Dim 标题1样式 As Style
    Dim 标题2样式 As Style
    Dim 标题3样式 As Style
    Dim 标题4样式 As Style
    
    Set 文档 = ActiveDocument
    
    '========================
    ' 处理：标题1
    '========================
    Set 标题1样式 = 文档.Styles("标题 1")
    
    '（1）将样式基础设为“无样式”
    ' 方式A：直接置空（多数版本可行）
    标题1样式.BaseStyle = ""
    ' 方式B（备选）：若A报错，可用占位样式过渡（用完可删除）
    'Dim 占位 As Style
    'Set 占位 = 文档.Styles.Add(Name:="无样式占位", Type:=wdStyleTypeParagraph)
    '标题1样式.BaseStyle = 占位
    '文档.Styles("无样式占位").Delete
    
    '（2）字体：中文=宋体；西文/数字=Times New Roman；加粗
    With 标题1样式.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 18    ' 小二号 = 18 pt
    End With
    
    '（3）段落：大纲级别=1级；左右缩进0；无首行缩进；段前0.5磅、段后0.5磅；1.5倍行距；居中
    With 标题1样式.ParagraphFormat
        .outlineLevel = wdOutlineLevel1
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0       ' 特殊缩进=无
        .SpaceBefore = 0.5
        .SpaceAfter = 0.5
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
        .CharacterUnitFirstLineIndent = 0
        .TabStops.ClearAll
    End With
    
    Call 全文重新套用样式(文档, 标题1样式)
    
    
    
    '========================
    ' 处理：标题2
    '========================
    Set 标题2样式 = 文档.Styles("标题 2")
    
    ' 样式基础：无样式
    标题2样式.BaseStyle = ""
    
    ' 字体：中文宋体，西文/数字 Times New Roman，加粗；四号=14pt
    With 标题2样式.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 14    ' 四号 = 14 pt
    End With
    
    ' 段落：大纲级别=2；左右缩进0；无首行缩进；段前0.5磅 段后0；1.5倍行距；左对齐
    With 标题2样式.ParagraphFormat
        .outlineLevel = wdOutlineLevel2
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = 0.5
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
        .CharacterUnitFirstLineIndent = 0
        .TabStops.ClearAll
    End With
    
    Call 全文重新套用样式(文档, 标题2样式)
    
    '========================
    ' 处理：标题3
    '========================
    Set 标题3样式 = 文档.Styles("标题 3")
    
    ' 样式基础：无样式
    标题3样式.BaseStyle = ""
    
    ' 字体：中文宋体，西文/数字 Times New Roman，加粗；小四号=12pt
    With 标题3样式.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 12    ' 小四号 = 12 pt
    End With
    
    ' 段落：大纲级别=3；左右缩进0；无首行缩进；段前0 段后0；1.5倍行距；左对齐
    With 标题3样式.ParagraphFormat
        .outlineLevel = wdOutlineLevel3
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
        .CharacterUnitFirstLineIndent = 0
        .TabStops.ClearAll
    End With
    
    Call 全文重新套用样式(文档, 标题3样式)
    
    '========================
    ' 处理：标题4
    '========================
    Set 标题4样式 = 文档.Styles("标题 4")
    
    ' 样式基础：无样式
    标题4样式.BaseStyle = ""
    
    ' 字体：中文宋体，西文/数字 Times New Roman，加粗；小四号=12pt
    With 标题4样式.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 12    ' 小四号 = 12 pt
    End With
    
    ' 段落：大纲级别=4；左右缩进0；无首行缩进；段前0 段后0；1.5倍行距；左对齐
    With 标题4样式.ParagraphFormat
        .outlineLevel = wdOutlineLevel4
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
        .CharacterUnitFirstLineIndent = 0
        .TabStops.ClearAll
    End With
    
    Call 全文重新套用样式(文档, 标题4样式)
    
    '========================
    ' 处理：表格标题（自定义样式“表格标题”）
    '========================
    Dim 表格标题样式 As Style
    
    ' 若不存在则创建为段落样式
    On Error Resume Next
    Set 表格标题样式 = 文档.Styles("表格标题")
    On Error GoTo 0
    If 表格标题样式 Is Nothing Then
        Set 表格标题样式 = 文档.Styles.Add(name:="表格标题", Type:=wdStyleTypeParagraph)
    End If
    
    ' 字体：中文=黑体；西文/数字=Times New Roman；五号=10.5 pt；加粗
    With 表格标题样式.Font
        .NameFarEast = "黑体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 10.5         ' 五号 = 10.5 pt
    End With
    
    ' 段落：行距1.5倍；段前0 段后0；左右缩进0；无首行缩进；清除制表位；非大纲标题
    With 表格标题样式.ParagraphFormat
        .outlineLevel = wdOutlineLevelBodyText
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphCenter
        .CharacterUnitFirstLineIndent = 0
        .TabStops.ClearAll
    End With
    
    ' ★ 设置完成后，立即让全文中已用“表格标题”样式的段落重新套用该样式
    Call 全文重新套用样式(文档, 表格标题样式)

    
    MsgBox "标题1～4样式已设置，并与多级列表模板4完成绑定。"
    MsgBox "标题1～4样式已按要求更新完成。"
End Sub


'——— 辅助：将文档中已应用某样式的段落，统一按该样式“重新套用”一次（无显式循环）
Private Sub 全文重新套用样式(ByVal 文档 As Document, ByVal 目标样式 As Style)
    With 文档.content.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .text = ""                         ' 不找文本，只按样式筛选
        .replacement.text = ""             ' 替换为“同名样式”
        .Style = 目标样式                  ' 查找：目标样式
        .replacement.Style = 目标样式      ' 替换：仍是目标样式（相当于重套样式）
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
