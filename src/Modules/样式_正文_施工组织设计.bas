Attribute VB_Name = "��ʽ_����_ʩ����֯���"
Sub һ���������ĸ�ʽ()
    Selection.Style = ActiveDocument.Styles("����")
    With ActiveDocument.Styles("����").Font
        .NameFarEast = "����"
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
    
    With ActiveDocument.Styles("����").ParagraphFormat
        .CharacterUnitFirstLineIndent = 2
        .outlineLevel = wdOutlineLevelBodyText
        .LeftIndent = CentimetersToPoints(0) ' ���0�ַ�
        .RightIndent = CentimetersToPoints(0) ' �Ҳ�0�ַ�
        .SpaceBefore = 0 ' ��ǰ0��
        .SpaceAfter = 0 ' �κ�0��
        .LineSpacingRule = wdLineSpace1pt5
    End With
    
End Sub
