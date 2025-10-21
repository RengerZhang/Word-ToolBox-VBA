Attribute VB_Name = "ͼƬ����_2_����ͼƬ��"
Option Explicit

'========================
' ���뵥��ͼƬ��ͼƬ���ݿؼ��棩
' 1) ��괦���� 1��2 ����ס�ͼƬ��λ��
' 2) ���������Ӧ
' 3) ��1�в��롾ͼƬ���ݿؼ������������ѡͼ���и߰� 4:3 Ԥ�����ӿռ�
' 4) ��2��д�롰ͼXXX  ���ڴ˴�¼��ͼ�����������ס�ͼƬ���⡱
'========================
Public Sub ���뵥��ͼƬ��_ͼƬ�ؼ���()
    '��һ����ʽ��
    Const STYLE_PIC_TABLE As String = "ͼƬ��λ��"
    Const PARA_STYLE_CAPTION As String = "ͼƬ����"
    Const CC_TAG As String = "PIC_4TO3"

    '������׼������ʽ����
    Dim doc As Document: Set doc = ActiveDocument
    EnsureTableStyleOnly doc, STYLE_PIC_TABLE

    '���������� 1��2 ��� �� ��ʽ �� ����Ӧ
    Dim rng As Range: Set rng = Selection.Range
    Dim tb As Table: Set tb = doc.Tables.Add(Range:=rng, NumRows:=2, NumColumns:=1)
    On Error Resume Next
    tb.Style = STYLE_PIC_TABLE
    On Error GoTo 0
    tb.AutoFitBehavior wdAutoFitWindow


    '���壩��1�У����롰ͼƬ���ݿؼ������������ѡͼ�Ի���
    Dim cc As ContentControl
    tb.cell(1, 1).Range.text = ""       ' ��յ�Ԫ������
    Set cc = doc.ContentControls.Add(wdContentControlPicture, tb.cell(1, 1).Range)

    With cc
        .Title = "ͼƬ��������룩"
        .tag = CC_TAG                   ' ���ں�����������
        .Appearance = wdContentControlBoundingBox
        ' ע��ͼƬ�ؼ���ռλ��ʾ��ͼ�ꣻҲ�ɸ���һ����ʾ�ı���
        .SetPlaceholderText , , "����˴�����ͼƬ�����������������ݣ���ѡ�д˰�ť��CTRL+Vճ����"
    End With
    ' ������ʾ�ؼ�
    cc.Range.ParagraphFormat.alignment = wdAlignParagraphCenter

    
    ' ��1�и߶��Զ���������ͼƬ��С����Ӧ
    With tb.rows(1)
        .HeightRule = wdRowHeightAuto     ' �и��Զ�
    End With


    '��������2�У���������ʽ
    With tb.cell(2, 1).Range
        .text = ����ͼƬ����ռλ_��������(ActiveDocument, tb.Range) & " ���ڴ˴�¼��ͼ��"
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = PARA_STYLE_CAPTION
        On Error GoTo 0
    End With


    '���ߣ����崹ֱ���У������ۣ�
    On Error Resume Next
    tb.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    On Error GoTo 0

    '���ˣ���λ������
    tb.Range.Select
    Selection.Collapse wdCollapseEnd
End Sub

'==============================
' ��ڣ����롰˫��ͼƬ��λ����ͼ�Ϸ���ͼ����һ��������ͼ����������ͼ����
'==============================
Public Sub ����˫��ͼƬ��_ͼƬ�ؼ���_˫��()
    Dim doc As Document: Set doc = ActiveDocument
    Dim tb As Table, r As Long, c As Long
    Dim cc As ContentControl
    Dim ��ͼռλ As String
    
    '��һ�������ɡ���ͼ��ռλ����������������̺�����
    ��ͼռλ = ����ͼƬ����ռλ_��������(doc, Selection.Range) & " ���ڴ˴�¼��ͼ��"
    
    '�������ڲ���㽨 3��2 ���
    Set tb = doc.Tables.Add(Selection.Range, 3, 2)
    
    '���������ñ����ʽ & ��������
    On Error Resume Next
    tb.Style = doc.Styles("ͼƬ��λ��")
    On Error GoTo 0
    
    With tb
        .AllowAutoFit = True
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100                       ' �������Ӧ����
        .rows.alignment = wdAlignRowCenter
        .rows.AllowBreakAcrossPages = False
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .Spacing = 0
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.SpaceAfter = 0
        .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    
    '���ģ���1�У��������ڷ�ͼ ���� �и��Զ����ڵ�Ԫ���ڲ��롰ͼƬ���ݿؼ���
    With tb.rows(1)
        .HeightRule = wdRowHeightAuto               ' �ؼ����Զ��иߣ���ͼ����Ӧ
    End With
    For c = 1 To 2
        With tb.cell(1, c).Range
            ' ��Ԫ����ʽ��ͼƬ��ʽ
            On Error Resume Next
            .Style = ActiveDocument.Styles("ͼƬ��ʽ")
            On Error GoTo 0
            ' ����/�������
            .ParagraphFormat.alignment = wdAlignParagraphCenter
            ' ���롰ͼƬ�����ݿؼ�
            Set cc = .ContentControls.Add(wdContentControlPicture)
            cc.Title = IIf(c = 1, "ͼƬa", "ͼƬb")
            cc.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
            cc.SetPlaceholderText , , "�����˴�����ͼƬ"
        End With
    Next c
    
    '���壩��2�У�������ͼ����a/b������ʽ=ͼƬ����
    Call ����ͼƬ����_��ͼ��ʽ
    
    With tb.rows(2)
        .HeightRule = wdRowHeightAtLeast            ' �ı��и�����С�߶ȸ���
        .Height = CentimetersToPoints(0.7)
    End With
    With tb.cell(2, 1).Range
        .text = "a�� ������ͼ��"
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = ActiveDocument.Styles("ͼƬ����-��ͼ")
        On Error GoTo 0
    End With
    With tb.cell(2, 2).Range
        .text = "b�� ������ͼ��"
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = ActiveDocument.Styles("ͼƬ����-��ͼ")
        On Error GoTo 0
    End With
    
    '��������3�У��ϲ�Ϊ��ͼ������ʽ=ͼƬ����
    tb.cell(3, 1).Merge tb.cell(3, 2)
    With tb.cell(3, 1).Range
        .text = ��ͼռλ
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = ActiveDocument.Styles("ͼƬ����")
        On Error GoTo 0
    End With
    
    '���ߣ��ѹ�����ڱ�󣬱��ڼ����༭
    tb.Range.Collapse wdCollapseEnd
    tb.Range.Select
End Sub


'========================
' ���ף�ȷ���������ʽ�����ڣ�������ۣ�
'========================
Private Sub EnsureTableStyleOnly(ByVal doc As Document, ByVal styleName As String)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    If Not st Is Nothing Then
        If st.Type <> wdStyleTypeTable Then
            st.Delete
            Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
        End If
    Else
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
    End If
End Sub

'==========================================================
' ����ͼƬ����ռλ���������°�����ͼ��
' A = �������(����H4��H3��H2��H1) �ı�ţ���ȡ������Ĭ�� "ͼ1.1-1"
' B = ���Ͻ�=���H3(�˻���H2/H1/����)���½�=��secStart��������H1/H2/H3(������ĩ)
'     ������ͳ�ơ�ͼƬ���⡱�����õ���� idx = n + 1
'==========================================================
Private Function ����ͼƬ����ռλ_��������(ByVal doc As Document, ByVal atRng As Range) As String
    Dim chapA As String
    Dim anchorB As Paragraph
    Dim chapB As String
    Dim secStart As Long, secEnd As Long
    Dim n As Long, idx As Long

    '��һ��A����������ţ�H4��H3��H2��H1��
    chapA = ȡ���������(atRng, Array(4, 3, 2, 1))
    If Len(chapA) = 0 Then
        ����ͼƬ����ռλ_�������� = "ͼ1.1-1"          ' ���δ��ʼ�� �� ����
        Exit Function
    End If

    '������B ���Ͻ磺���H3 �� ����H2 �� ����H1 �� ��������
    Set anchorB = ȡ����������(atRng, Array(3, 2, 1))
    If anchorB Is Nothing Then
        secStart = doc.content.Start
        chapB = ""                                  ' �����ޱ��
    Else
        secStart = anchorB.Range.End                ' �Ͻ��� End���ų������У�
        chapB = ��ȫȡ�б���(anchorB)
    End If

    '������B ���½磺�� secStart ����� H1/H2/H3 ���������ȳ����ߣ�������ĩ
    secEnd = �����½�������ֵ�(doc, secStart)

    '���ģ�ͳ������ [secStart..secEnd) �ڵġ�ͼƬ���⡱�������ϸ�/���׶�ѡһ��
    n = ͳ������ͼƬ������(doc, secStart, secEnd, chapB)

    '���壩��װռλ
    idx = n + 1
    ����ͼƬ����ռλ_�������� = "ͼ" & chapA & "-" & CStr(idx)
End Function

'==========================================================
' ȡ��������ϸ������𼯺ϡ��ı����ţ����� levels = Array(4,3,2,1)��
' �ҵ������ض���� ListString���Ҳ������� ""
'==========================================================
Private Function ȡ���������(ByVal atRng As Range, ByVal levels As Variant) As String
    Dim p As Paragraph
    Set p = ȡ����������(atRng, levels)
    If p Is Nothing Then
        ȡ��������� = ""
    Else
        ȡ��������� = ��ȫȡ�б���(p)
    End If
End Function

'==========================================================
' ȡ������⡰������󡱣���ǰ���ݣ�levels ����Array(3,2,1)��
'==========================================================
Private Function ȡ����������(ByVal atRng As Range, ByVal levels As Variant) As Paragraph
    Dim p As Paragraph, prev As Paragraph, i As Long
    Set p = atRng.Paragraphs(1)

    Do While Not p Is Nothing
        For i = LBound(levels) To UBound(levels)
            If �����Ƿ�ָ���������(p, CLng(levels(i))) Then
                Set ȡ���������� = p
                Exit Function
            End If
        Next i
        Set p = ��һ������(p)
    Loop
    Set ȡ���������� = Nothing
End Function

'==========================================================
' �ж϶����Ƿ�ָ��������⣨�������ġ����� n����Ӣ�ġ�Heading n����
'==========================================================
Private Function �����Ƿ�ָ���������(ByVal p As Paragraph, ByVal lvl As Long) As Boolean
    On Error Resume Next
    Dim nm As String
    If TypeName(p.Range.Style) = "Style" Then
        nm = p.Range.Style.nameLocal
    Else
        nm = CStr(p.Range.Style)
    End If
    On Error GoTo 0

    nm = LCase$(nm)
    �����Ƿ�ָ��������� = (nm = LCase$("���� " & lvl)) Or (nm = LCase$("heading " & lvl))
End Function

'==========================================================
' ��һ�����䣨��ȫȡ����
'==========================================================
'��������һ�����䣨Collapse + Move/Expand������ Duplicate��
Private Function ��һ������(ByVal p As Paragraph) As Paragraph
    Dim r As Range
    Set r = p.Range.Duplicate ' �����Ե��� Duplicate�����пɸ�Ϊ��Set r = p.Range

    r.Collapse wdCollapseStart
    If r.Start = 0 Then Exit Function               ' ����
    r.MoveStart wdCharacter, -1                     ' ��ǰ�� 1 ���ַ�
    r.Expand wdParagraph                            ' ��չΪ��һ������
    Set ��һ������ = r.Paragraphs(1)
End Function



'==========================================================
' ��ȫȡ�б��ţ�ListString ����ȡ���� �� ���� ""��
'==========================================================
Private Function ��ȫȡ�б���(ByVal p As Paragraph) As String
    On Error Resume Next
    ��ȫȡ�б��� = Trim$(p.Range.ListFormat.ListString)
    If Err.Number <> 0 Then ��ȫȡ�б��� = "": Err.Clear
    On Error GoTo 0
End Function

'==========================================================
' �����½磺�� secStart ���� H1/H2/H3��ȡ���ȳ����ߵ� Start�����ޡ���ĩ
'==========================================================
Private Function �����½�������ֵ�(ByVal doc As Document, ByVal secStart As Long) As Long
    Dim p1 As Long, p2 As Long, p3 As Long
    p1 = ������һ���������(doc, secStart, 1)
    p2 = ������һ���������(doc, secStart, 2)
    p3 = ������һ���������(doc, secStart, 3)

    Dim m As Long: m = ��С����(p1, p2, p3)
    If m = -1 Then
        �����½�������ֵ� = doc.content.End
    Else
        �����½�������ֵ� = m
    End If
End Function

'==========================================================
' ���Ҵ� pos �����һ�� ����lvl������㣻�Ҳ������� -1
'==========================================================
Private Function ������һ���������(ByVal doc As Document, ByVal pos As Long, ByVal lvl As Long) As Long
    Dim rng As Range: Set rng = doc.Range(Start:=pos, End:=doc.content.End)
    With rng.Find
        .ClearFormatting
        On Error Resume Next
        .Style = doc.Styles(IIf(lvl = 1, "���� 1", IIf(lvl = 2, "���� 2", "���� 3")))
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
    End With
    If rng.Find.Execute Then
        ������һ��������� = rng.Start
        Exit Function
    End If
    ' Ӣ����ʽ������
    Set rng = doc.Range(Start:=pos, End:=doc.content.End)
    With rng.Find
        .ClearFormatting
        On Error Resume Next
        .Style = doc.Styles(IIf(lvl = 1, "Heading 1", IIf(lvl = 2, "Heading 2", "Heading 3")))
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
    End With
    If rng.Find.Execute Then
        ������һ��������� = rng.Start
    Else
        ������һ��������� = -1
    End If
End Function

'==========================================================
' ��������λ�õġ���С��ֵ������ȫΪ -1�������ã����� -1
'==========================================================
Private Function ��С����(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim t As Variant: t = Array(a, b, c)
    Dim i As Long, best As Long: best = -1
    For i = LBound(t) To UBound(t)
        If CLng(t(i)) >= 0 Then
            If best = -1 Or CLng(t(i)) < best Then best = CLng(t(i))
        End If
    Next i
    ��С���� = best
End Function

'==========================================================
' ͳ�������ڡ�ͼƬ���⡱���������ϸ�ģʽ���޶�ǰ׺���� chapB="" ����ʽ������
'==========================================================
Private Function ͳ������ͼƬ������(ByVal doc As Document, ByVal startPos As Long, ByVal endPos As Long, ByVal chapB As String) As Long
    Dim scan As Range: Set scan = doc.Range(Start:=startPos, End:=endPos)
    Dim p As Paragraph, t As String, n As Long
    For Each p In scan.Paragraphs
        If ������ʽ����(p, "ͼƬ����") Then
            t = ����ɼ��ı�(p.Range.text)
            If Len(chapB) = 0 Then
                n = n + 1
            Else
                ' ƥ�䣺^ͼ<chapB>- �� ^ͼ<chapB>.
                If Left$(t, Len("ͼ" & chapB & "-")) = "ͼ" & chapB & "-" _
                   Or Left$(t, Len("ͼ" & chapB & ".")) = "ͼ" & chapB & "." Then
                    n = n + 1
                End If
            End If
        End If
    Next
    ͳ������ͼƬ������ = n
End Function

Private Function ������ʽ����(ByVal p As Paragraph, ByVal styleName As String) As Boolean
    On Error Resume Next
    Dim nm As String
    If TypeName(p.Range.Style) = "Style" Then
        nm = p.Range.Style.nameLocal
    Else
        nm = CStr(p.Range.Style)
    End If
    On Error GoTo 0
    ������ʽ���� = (LCase$(nm) = LCase$(styleName))
End Function

' ȥ���س�/��Ԫ�������/�հ�
Private Function ����ɼ��ı�(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")
    ����ɼ��ı� = Trim$(s)
End Function

