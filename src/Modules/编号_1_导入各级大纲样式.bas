Attribute VB_Name = "���_1_������������ʽ"
Option Explicit

'��һ�������ʽ��������ģ�鼶����ȫģ���κι��̸��ã�
Private Const STYLE_TABLE_NORMAL As String = "��׼�����ʽ"   ' ��ͨ������Ϊ��ǣ��������
Private Const STYLE_TABLE_PIC    As String = "ͼƬ��λ��"     ' ͼƬ���޿��� + 0 �ڱ߾�

Sub һ������ȫ����ʽ()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Call һ���������ĸ�ʽ
    Call ���ñ�����ʽ(doc, "���� 1", 18, wdOutlineLevel1, 0.5, 0.5, wdAlignParagraphCenter)
    Call ���ñ�����ʽ(doc, "���� 2", 14, wdOutlineLevel2, 0.5, 0, wdAlignParagraphLeft)
    Call ���ñ�����ʽ(doc, "���� 3", 12, wdOutlineLevel3, 0, 0, wdAlignParagraphLeft)
    Call ���ñ�����ʽ(doc, "���� 4", 12, wdOutlineLevel4, 0, 0, wdAlignParagraphLeft)
    Call ������������ʽ(doc, "����ʽ��1����", wdOutlineLevelBodyText)
    Call ������������ʽ(doc, "����ʽ����1����", wdOutlineLevelBodyText)
    Call ������������ʽ(doc, "����ʽ���١�", wdOutlineLevelBodyText)
    Call EnsureStandardTableStyle
    Call ����ͼƬ����_��ͼ��ʽ
    
    '========================
    ' 4. ������ / ͼƬ����
    '========================
    Dim styleTableCaption As Style, stylePicCaption As Style, stylePicPara As Style   ' �� ���� stylePicPara
    
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
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpace1pt5   '1.5���о�
        .SpaceBefore = Application.LinesToPoints(0)   ' ��ǰ 0 ��
        .SpaceAfter = Application.LinesToPoints(0)    ' �κ� 0 ��
        ' ===�����������¶�ͬҳ===
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
    With stylePicCaption.ParagraphFormat
        .KeepWithNext = False
        .outlineLevel = wdOutlineLevelBodyText
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphCenter
        .LineSpacingRule = wdLineSpace1pt5   '1.5���о�
        .SpaceBefore = Application.LinesToPoints(0)   ' ��ǰ 0 ��
        .SpaceAfter = Application.LinesToPoints(0)    ' �κ� 0 ��
    End With
    
    '========================
    ' 4+. ͼƬ������ʽ��ͼƬ��ʽ��������ʽ��
    '========================
    On Error Resume Next
    Set stylePicPara = doc.Styles("ͼƬ��ʽ")
    If stylePicPara Is Nothing Then
        Set stylePicPara = doc.Styles.Add(name:="ͼƬ��ʽ", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    
    ' ������ʽ�����̳��κ���ʽ
    stylePicPara.BaseStyle = ""
    stylePicPara.NextParagraphStyle = doc.Styles("����")
    
    ' ��������Ҫ������ԣ����ౣ��Ĭ��
    With stylePicPara.ParagraphFormat
        .outlineLevel = wdOutlineLevelBodyText     ' ���ļ�
        .LeftIndent = 0                            ' ������
        .RightIndent = 0
        .FirstLineIndent = 0                       ' ����������
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphCenter        ' ����
        .KeepWithNext = True                       ' ���¶�ͬҳ
        ' �������о�/��ǰ�κ�/�߿����/�Ʊ�λ�ȣ�����Ĭ��
    End With
    
    '========================
    ' 5. �����ʽ
    '========================
    Dim tblStyleStd As Style, tblStylePic As Style
    
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
    
    
    '========================
    ' TOC 1~3�����������ص� + �����оࣨ������ã������Ժ����������
    '========================
    
    '��һ��TOC 1
    With doc.Styles(wdStyleTOC1).ParagraphFormat
        .FirstLineIndent = 0                 ' �� ������������������
        .CharacterUnitFirstLineIndent = 0    ' �� �����������ַ�������
        .LineSpacingRule = wdLineSpaceSingle ' �� �о���Ϊ��������
    End With
    
    '������TOC 2
    With doc.Styles(wdStyleTOC2).ParagraphFormat
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
    
    '������TOC 3
    With doc.Styles(wdStyleTOC3).ParagraphFormat
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
    
    '����TOC ����1�����������ص� + �����о� + ����/�Ӵ�/С��
    With ActiveDocument.Styles("TOC ����")
        '��һ�����䣺�������������о�=����
        With .ParagraphFormat
            .FirstLineIndent = 0                   ' �� ������������������
            .CharacterUnitFirstLineIndent = 0      ' �� �����������ַ�������
            .LineSpacingRule = wdLineSpaceSingle   ' �� �о���Ϊ��������
            .alignment = wdAlignParagraphCenter
        End With
        '���������壺���� + �Ӵ� + С�ģ�12pt��
        With .Font
            .NameFarEast = "����"                  ' �� ��������=����
            .Size = 18                             ' �� С�ĺ�=12 ��
            .bold = True                           ' �� �Ӵ�
            .Color = wdColorBlack
            .NameAscii = "Times New Roman"
        End With
    End With
    
    
    
    MsgBox "������ʽһ��������ɣ�"
End Sub
'========================
' �������������ñ�����ʽ���ɴ���ǰ/�κ��С������뷽ʽ��
' ������
'   styleName   ������ʽ�����硰���� 1����
'   fontSize    �ֺţ�pt��
'   olLevel     ��ټ���
'   beforeLines ��ǰ(��)  ���� ��ѡ��Ĭ��0
'   afterLines  �κ�(��)  ���� ��ѡ��Ĭ��0
'   align       ���뷽ʽ  ���� ��ѡ��Ĭ�������
'========================
Private Sub ���ñ�����ʽ( _
    doc As Document, ByVal styleName As String, ByVal fontSize As Single, _
    ByVal olLevel As WdOutlineLevel, _
    Optional ByVal beforeLines As Single = 0, _
    Optional ByVal afterLines As Single = 0, _
    Optional ByVal align As WdParagraphAlignment = wdAlignParagraphLeft)

    Dim st As Style
    Set st = doc.Styles(styleName)   ' �ٶ����ñ�����ʽ�Ѵ���

    '��һ������
    With st.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .bold = True
        .Size = fontSize
    End With

    '����������
    With st.ParagraphFormat
        .outlineLevel = olLevel
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = align
        ' �����С�����Ϊ��
'        .SpaceBefore = Application.LinesToPoints(beforeLines)
'        .SpaceAfter = Application.LinesToPoints(afterLines)
        .SpaceBefore = beforeLines
        .SpaceAfter = afterLines
    End With

    '��������һ����ʽ�ص�����
    st.NextParagraphStyle = doc.Styles("����")
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


Public Sub �������ֱ����ʽ_������()
    Dim doc As Document: Set doc = ActiveDocument
    Dim styStd As Style, styPic As Style

    '��һ��ȷ�����ڡ���׼�����ʽ������ͨ��ı���ã������κ��Ӿ����ã�
    Set styStd = EnsureTableStyleOnly(doc, STYLE_TABLE_NORMAL)
    ' ���������� .Table/.ParagraphFormat/.Font����ȫ����Ĭ�ϣ��ɺ�����ʽ������ͳһ������

    '������ȷ�����ڡ�ͼƬ��λ����ͼƬ��ר�ã�
    Set styPic = EnsureTableStyleOnly(doc, STYLE_TABLE_PIC)
    With styPic.Table
        '��1���ر�ȫ������
        .Borders.enable = False
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone

        '��2���и߱��֡�����Ӧ����Ĭ�ϼ�Ϊ�Զ��иߣ����ﲻ��ʽ���ã�

        '��3����Ԫ���ĸ�����ľ����Ϊ 0
        .TopPadding = 0
        .BottomPadding = 0
        .LeftPadding = 0
        .RightPadding = 0

        '��4�������ö���/�ַ�������ԣ�ˮƽ/��ֱ���е������������̴���
    End With
End Sub
' === ��׼�������ʽ�����������޸ģ��������򴴽���������ʽ��===
Public Sub EnsureStandardTableStyle()
    Dim stdStyle As Style

    ' 1) ���Ի�ȡ�Ѵ��ڵ���ʽ
    On Error Resume Next
    Set stdStyle = ActiveDocument.Styles("��׼�������ʽ")
    On Error GoTo 0

    ' 2) �������򴴽������ڵ����ͷǡ�������ʽ�����ؽ�Ϊ������ʽ
    If stdStyle Is Nothing Then
        Set stdStyle = ActiveDocument.Styles.Add( _
            name:="��׼�������ʽ", Type:=wdStyleTypeParagraph)
    ElseIf stdStyle.Type <> wdStyleTypeParagraph Then
        stdStyle.Delete
        Set stdStyle = ActiveDocument.Styles.Add( _
            name:="��׼�������ʽ", Type:=wdStyleTypeParagraph)
    End If

    ' 3) ������ʽ����
    With stdStyle
        .AutomaticallyUpdate = False

        ' ������Ϊ��������ʽ������������������˵� Normal
        On Error Resume Next
        .BaseStyle = ""                                ' ���̳�����
        If Err.Number <> 0 Then
            Err.Clear
            .BaseStyle = ActiveDocument.Styles(wdStyleNormal)
        End If
        On Error GoTo 0

        ' ====== ���� ======
        With .Font
            .NameFarEast = "����"                      ' ����
            .NameAscii = "Times New Roman"            ' Ӣ��/����
            .Size = 10.5                               ' ���
            .bold = False
            .Color = wdColorAutomatic
        End With

        ' ====== ���� ======
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpaceSingle       ' �����о�
            .SpaceBefore = 0                           ' ��ǰ 0
            .SpaceAfter = 0                            ' �κ� 0
            .LeftIndent = 0                            ' ������
            .RightIndent = 0                           ' ������
            .FirstLineIndent = 0                       ' ��������
            .CharacterUnitFirstLineIndent = 0
            .alignment = wdAlignParagraphCenter        ' ����
            '.NoSpaceBetweenParagraphsOfSameStyle = True ' �ɰ�������
        End With
    End With
End Sub
'========================================================
' ���ߣ�����֤�������ʽ���Ĵ�����������ȷ�������κζ���/�ַ����ԣ�
' - ��ͬ����ʽ���ڵ����Ͳ��ǡ������ʽ������ɾ�����ԡ������ʽ���ؽ�
'========================================================
Private Function EnsureTableStyleOnly(ByVal doc As Document, ByVal styleName As String) As Style
    Dim st As Style

    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0

    If Not st Is Nothing Then
        If st.Type <> wdStyleTypeTable Then
            '��һ��ͬ�������Ͳ��ԣ�ɾ�����ؽ�Ϊ�������ʽ��
            st.Delete
            Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
        End If
    Else
        '���������������½�
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
    End If

    Set EnsureTableStyleOnly = st
End Function
'==============================
'��������ͼƬ��ͼ����ʽ�������ڡ�ͼƬ���⡱��
'  Ĭ�ϣ����塢12pt���ǼӴ֡������оࡢ���С��޼̳С��ض��뵽����
'==============================
Public Sub ����ͼƬ����_��ͼ��ʽ()
    '��һ������/������ʽ����
    Call EnsureParagraphStyle( _
        styleName:="ͼƬ����-��ͼ", _
        nameCN:="����", nameEN:="Times New Roman", _
        ptSize:=10.5, isBold:=False, _
        lineRule:=wdLineSpace1pt5, align:=wdAlignParagraphCenter)

    '����������ͨ�����ԣ��޼̳С��رա����뵽���񡱡���ǰ��/��������
    With ActiveDocument.Styles("ͼƬ����-��ͼ")
        .AutomaticallyUpdate = False
        On Error Resume Next
        .BaseStyle = "ͼƬ����"                           ' �������κ���ʽ���޼̳У�
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
'  ��һ���Ȱ�ȫ��ȡ��ʽ���󣨲����ȡʧ�ܣ�
'  �������������жϣ��� Nothing �� �½���������������
'  �������������������/��������
'===============================
Public Sub EnsureParagraphStyle( _
    ByVal styleName As String, _
    ByVal nameCN As String, ByVal nameEN As String, _
    ByVal ptSize As Single, ByVal isBold As Boolean, _
    ByVal lineRule As WdLineSpacing, ByVal align As WdParagraphAlignment)

    Dim st As Style

    '��һ����ȫ��ȡ��ʽ����
    On Error Resume Next
    Set st = ActiveDocument.Styles(styleName)
    If Err.Number <> 0 Then
        Err.Clear
        Set st = Nothing
    End If
    On Error GoTo 0

    '������������ or ���Ͳ��� �� �½�Ϊ��������ʽ��
    If st Is Nothing Then
        Set st = ActiveDocument.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    Else
        ' ע�⣺VBA �޶�·��ֵ�������֧���ٷ��� st.Type
        If st.Type <> wdStyleTypeParagraph Then
            On Error Resume Next
            st.Delete                      ' ͬ�������Ͳ��ԣ�����/�ַ���ʽ����ɾ���ؽ�
            On Error GoTo 0
            Set st = ActiveDocument.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
        End If
    End If

    '������ͳһ������ʽ����
    With st
        .AutomaticallyUpdate = False

        ' ������ʽ����ָ������/���桱����ͬ�����ݴ�
        On Error Resume Next
        .BaseStyle = ""                 ' �� �ؼ����������κ���ʽ
        Err.Clear
        On Error GoTo 0

        ' ���壨��Ӣ�ķֱ����ã�
        With .Font
            .NameFarEast = nameCN
            .NameAscii = nameEN
            .Size = ptSize
            .bold = isBold
            .Color = wdColorAutomatic
        End With

        ' ���䣺���� + 1.5���оࣨ��ö�� wdLineSpace1pt5��
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

