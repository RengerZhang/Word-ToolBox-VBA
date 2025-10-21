Attribute VB_Name = "��ʽ_ʩ����֯��ʽ_һ������"
Sub һ������ȫ����ʽ2()
    Dim doc As Document
    Set doc = ActiveDocument
    
    '========================
    ' 1. ������ʽ
    '========================
    һ���������ĸ�ʽ
    
    '========================
    ' 2. ����1��4
    '========================
    Call ���ñ�����ʽ(doc, "���� 1", 18, wdOutlineLevel1)
    Call ���ñ�����ʽ(doc, "���� 2", 14, wdOutlineLevel2)
    Call ���ñ�����ʽ(doc, "���� 3", 12, wdOutlineLevel3)
    Call ���ñ�����ʽ(doc, "���� 4", 12, wdOutlineLevel4)
    Call ������������ʽ(doc, "����ʽ��1����", wdOutlineLevelBodyText)
    Call ������������ʽ(doc, "����ʽ����1����", wdOutlineLevelBodyText)
    Call ������������ʽ(doc, "����ʽ���١�", wdOutlineLevelBodyText)
    
    '========================
    ' 4. ������ / ͼƬ����
    '========================
    Dim styleTableCaption As Style, stylePicCaption As Style
    On Error Resume Next
    Set styleTableCaption = doc.Styles("������")
    If styleTableCaption Is Nothing Then
        Set styleTableCaption = doc.Styles.Add(name:="������", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    With styleTableCaption.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .bold = True
        .Size = 10.5
    End With
    With styleTableCaption.ParagraphFormat
        .outlineLevel = wdOutlineLevelBodyText
        .alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpace1pt5
        .FirstLineIndent = 0 ' ��������
        .CharacterUnitFirstLineIndent = 0
        .KeepWithNext = True
        
    End With
    
    ' ͼƬ���� = �̳б�����
    On Error Resume Next
    Set stylePicCaption = doc.Styles("ͼƬ����")
    If stylePicCaption Is Nothing Then
        Set stylePicCaption = doc.Styles.Add(name:="ͼƬ����", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    stylePicCaption.BaseStyle = "������"
    
    '========================
    ' 5. �����ʽ
    '========================
    Dim tblStyleStd As Style, tblStylePic As Style
    
    ' ��׼�����ʽ
    On Error Resume Next
    Set tblStyleStd = doc.Styles("��׼�����ʽ")
    If tblStyleStd Is Nothing Then
        Set tblStyleStd = doc.Styles.Add(name:="��׼�����ʽ", Type:=wdStyleTypeTable)
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
        ' �ڿ��߱���Ĭ��
    End With
    
    ' ͼƬ��λ��
    On Error Resume Next
    Set tblStylePic = doc.Styles("ͼƬ��λ��")
    If tblStylePic Is Nothing Then
        Set tblStylePic = doc.Styles.Add(name:="ͼƬ��λ��", Type:=wdStyleTypeTable)
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
    
    MsgBox "���ġ�����1��7����������/ͼƬ���⡢�����ʽ��һ��������ɣ�"
End Sub

'========================
' �������������ñ�����ʽ
'========================
Private Sub ���ñ�����ʽ(doc As Document, ByVal styleName As String, ByVal fontSize As Single, ByVal olLevel As WdOutlineLevel)
    Dim st As Style
    Set st = doc.Styles(styleName)
    st.BaseStyle = ""
    With st.Font
        .NameFarEast = "����"
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
    
    st.NextParagraphStyle = ActiveDocument.Styles("����")
End Sub

'========================
' ����������������������ʽ
'========================
Private Sub ������������ʽ(doc As Document, ByVal styleName As String, ByVal olLevel As WdOutlineLevel)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    If st Is Nothing Then
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    
    st.BaseStyle = "����"
    With st.ParagraphFormat
        .outlineLevel = olLevel
        .CharacterUnitFirstLineIndent = 2
    End With
    
    st.NextParagraphStyle = ActiveDocument.Styles("����")
End Sub
Private Sub һ���������ĸ�ʽ()
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
        .alignment = wdAlignParagraphLeft
    End With
    
End Sub

