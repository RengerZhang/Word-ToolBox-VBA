Attribute VB_Name = "��ʽ_�ļ�����_������"
Private Sub ��ȡ��ǰ�༶�б�ģ��()
    Dim oDoc As Document
    Set oDoc = Word.ActiveDocument
    Dim oRng As Range
    Dim oList As List
    Dim oListFormat As ListFormat
    Dim oP As Paragraph
    Set oRng = Word.Selection.Range
    With oRng
        '��ȡ��ǰѡ���������ڵĵ�һ���б���Ŀ��ŵ��ַ���,����"2.3.1"
       MsgBox .ListFormat.ListType
    End With
End Sub

Sub ����ǰ�ļ�����������()
    Dim �ĵ� As Document
    Dim ����1��ʽ As Style
    Dim ����2��ʽ As Style
    Dim ����3��ʽ As Style
    Dim ����4��ʽ As Style
    
    Set �ĵ� = ActiveDocument
    
    '========================
    ' ��������1
    '========================
    Set ����1��ʽ = �ĵ�.Styles("���� 1")
    
    '��1������ʽ������Ϊ������ʽ��
    ' ��ʽA��ֱ���ÿգ������汾���У�
    ����1��ʽ.BaseStyle = ""
    ' ��ʽB����ѡ������A��������ռλ��ʽ���ɣ������ɾ����
    'Dim ռλ As Style
    'Set ռλ = �ĵ�.Styles.Add(Name:="����ʽռλ", Type:=wdStyleTypeParagraph)
    '����1��ʽ.BaseStyle = ռλ
    '�ĵ�.Styles("����ʽռλ").Delete
    
    '��2�����壺����=���壻����/����=Times New Roman���Ӵ�
    With ����1��ʽ.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 18    ' С���� = 18 pt
    End With
    
    '��3�����䣺��ټ���=1������������0����������������ǰ0.5�����κ�0.5����1.5���оࣻ����
    With ����1��ʽ.ParagraphFormat
        .outlineLevel = wdOutlineLevel1
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0       ' ��������=��
        .SpaceBefore = 0.5
        .SpaceAfter = 0.5
        .LineSpacingRule = wdLineSpace1pt5
        .alignment = wdAlignParagraphLeft
        .CharacterUnitFirstLineIndent = 0
        .TabStops.ClearAll
    End With
    
    Call ȫ������������ʽ(�ĵ�, ����1��ʽ)
    
    
    
    '========================
    ' ��������2
    '========================
    Set ����2��ʽ = �ĵ�.Styles("���� 2")
    
    ' ��ʽ����������ʽ
    ����2��ʽ.BaseStyle = ""
    
    ' ���壺�������壬����/���� Times New Roman���Ӵ֣��ĺ�=14pt
    With ����2��ʽ.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 14    ' �ĺ� = 14 pt
    End With
    
    ' ���䣺��ټ���=2����������0����������������ǰ0.5�� �κ�0��1.5���оࣻ�����
    With ����2��ʽ.ParagraphFormat
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
    
    Call ȫ������������ʽ(�ĵ�, ����2��ʽ)
    
    '========================
    ' ��������3
    '========================
    Set ����3��ʽ = �ĵ�.Styles("���� 3")
    
    ' ��ʽ����������ʽ
    ����3��ʽ.BaseStyle = ""
    
    ' ���壺�������壬����/���� Times New Roman���Ӵ֣�С�ĺ�=12pt
    With ����3��ʽ.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 12    ' С�ĺ� = 12 pt
    End With
    
    ' ���䣺��ټ���=3����������0����������������ǰ0 �κ�0��1.5���оࣻ�����
    With ����3��ʽ.ParagraphFormat
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
    
    Call ȫ������������ʽ(�ĵ�, ����3��ʽ)
    
    '========================
    ' ��������4
    '========================
    Set ����4��ʽ = �ĵ�.Styles("���� 4")
    
    ' ��ʽ����������ʽ
    ����4��ʽ.BaseStyle = ""
    
    ' ���壺�������壬����/���� Times New Roman���Ӵ֣�С�ĺ�=12pt
    With ����4��ʽ.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 12    ' С�ĺ� = 12 pt
    End With
    
    ' ���䣺��ټ���=4����������0����������������ǰ0 �κ�0��1.5���оࣻ�����
    With ����4��ʽ.ParagraphFormat
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
    
    Call ȫ������������ʽ(�ĵ�, ����4��ʽ)
    
    '========================
    ' ���������⣨�Զ�����ʽ�������⡱��
    '========================
    Dim ��������ʽ As Style
    
    ' ���������򴴽�Ϊ������ʽ
    On Error Resume Next
    Set ��������ʽ = �ĵ�.Styles("������")
    On Error GoTo 0
    If ��������ʽ Is Nothing Then
        Set ��������ʽ = �ĵ�.Styles.Add(name:="������", Type:=wdStyleTypeParagraph)
    End If
    
    ' ���壺����=���壻����/����=Times New Roman�����=10.5 pt���Ӵ�
    With ��������ʽ.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = True
        .Size = 10.5         ' ��� = 10.5 pt
    End With
    
    ' ���䣺�о�1.5������ǰ0 �κ�0����������0������������������Ʊ�λ���Ǵ�ٱ���
    With ��������ʽ.ParagraphFormat
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
    
    ' �� ������ɺ�������ȫ�������á������⡱��ʽ�Ķ����������ø���ʽ
    Call ȫ������������ʽ(�ĵ�, ��������ʽ)

    
    MsgBox "����1��4��ʽ�����ã�����༶�б�ģ��4��ɰ󶨡�"
    MsgBox "����1��4��ʽ�Ѱ�Ҫ�������ɡ�"
End Sub


'������ ���������ĵ�����Ӧ��ĳ��ʽ�Ķ��䣬ͳһ������ʽ���������á�һ�Σ�����ʽѭ����
Private Sub ȫ������������ʽ(ByVal �ĵ� As Document, ByVal Ŀ����ʽ As Style)
    With �ĵ�.content.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .text = ""                         ' �����ı���ֻ����ʽɸѡ
        .replacement.text = ""             ' �滻Ϊ��ͬ����ʽ��
        .Style = Ŀ����ʽ                  ' ���ң�Ŀ����ʽ
        .replacement.Style = Ŀ����ʽ      ' �滻������Ŀ����ʽ���൱��������ʽ��
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
