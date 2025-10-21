Attribute VB_Name = "样式_施工组织样式_一键导入"
Sub 一键设置全部样式2()
    Dim doc As Document
    Set doc = ActiveDocument
    
    '========================
    ' 1. 正文样式
    '========================
    一键设置正文格式
    
    '========================
    ' 2. 标题1～4
    '========================
    Call 设置标题样式(doc, "标题 1", 18, wdOutlineLevel1)
    Call 设置标题样式(doc, "标题 2", 14, wdOutlineLevel2)
    Call 设置标题样式(doc, "标题 3", 12, wdOutlineLevel3)
    Call 设置标题样式(doc, "标题 4", 12, wdOutlineLevel4)
    Call 创建条款项样式(doc, "条样式【1）】", wdOutlineLevelBodyText)
    Call 创建条款项样式(doc, "款样式【（1）】", wdOutlineLevelBodyText)
    Call 创建条款项样式(doc, "项样式【①】", wdOutlineLevelBodyText)
    
    '========================
    ' 4. 表格标题 / 图片标题
    '========================
    Dim styleTableCaption As Style, stylePicCaption As Style
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
        .alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpace1pt5
        .FirstLineIndent = 0 ' 首行缩进
        .CharacterUnitFirstLineIndent = 0
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
    
    '========================
    ' 5. 表格样式
    '========================
    Dim tblStyleStd As Style, tblStylePic As Style
    
    ' 标准表格样式
    On Error Resume Next
    Set tblStyleStd = doc.Styles("标准表格样式")
    If tblStyleStd Is Nothing Then
        Set tblStyleStd = doc.Styles.Add(name:="标准表格样式", Type:=wdStyleTypeTable)
    End If
    On Error GoTo 0
    With tblStyleStd.Table
        .Borders.enable = True
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineWidth = wdLineWidth150pt
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderLeft).LineWidth = wdLineWidth150pt
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        ' 内框线保留默认
    End With
    
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
    
    MsgBox "正文、标题1～7、条款项、表格/图片标题、表格样式已一键设置完成！"
End Sub

'========================
' 辅助函数：设置标题样式
'========================
Private Sub 设置标题样式(doc As Document, ByVal styleName As String, ByVal fontSize As Single, ByVal olLevel As WdOutlineLevel)
    Dim st As Style
    Set st = doc.Styles(styleName)
    st.BaseStyle = ""
    With st.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .bold = True
        .Size = fontSize
    End With
    With st.ParagraphFormat
        .outlineLevel = olLevel
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
    End With
    
    st.NextParagraphStyle = ActiveDocument.Styles("正文")
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

