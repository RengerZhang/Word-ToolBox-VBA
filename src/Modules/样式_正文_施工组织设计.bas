Attribute VB_Name = "样式_正文_施工组织设计"
Sub 一键设置正文格式()
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
    End With
    
End Sub
