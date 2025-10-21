Attribute VB_Name = "������_6_�����Լ�"
Option Explicit

'������ڣ������Լ죨�����ĵ���
Public Sub �Լ�_������_����()
    Dim doc As Document: Set doc = ActiveDocument
    Dim a() As Long: a = BuildPrevStartArray(doc)

    With New �Լ챨��
        .LoadReportLite doc.Tables.Count, a   ' �� ֻ�������� + ÿ��ǰ��Start
        .Show vbModeless
    End With
End Sub

'����Ϊÿ������㡰�ͽ���һ�ǿնΡ��� Start���Ҳ���=-1��
Private Function BuildPrevStartArray(ByVal doc As Document) As Long()
    Dim n As Long: n = doc.Tables.Count
    Dim arr() As Long: ReDim arr(1 To n)
    Dim i As Long, p As Paragraph

    For i = 1 To n
        Set p = PrevNonEmptyParaForTable(doc.Tables(i))
        If p Is Nothing Then
            arr(i) = -1
        Else
            arr(i) = p.Range.Start
        End If
        If (i And 31) = 0 Then DoEvents  ' ������
    Next
    BuildPrevStartArray = arr
End Function

'�����ͽ���һ�ǿնΣ��������հ׶Σ�
Private Function PrevNonEmptyParaForTable(ByVal tbl As Table) As Paragraph
    Dim r As Range: Set r = tbl.Range: r.Collapse wdCollapseStart
    Dim p As Paragraph: Set p = r.Paragraphs(1).Previous
    Do While Not p Is Nothing
        If Len(TrimVisible(p.Range.text)) > 0 Then
            Set PrevNonEmptyParaForTable = p
            Exit Function
        End If
        Set p = p.Previous
    Loop
End Function

'��������ɼ��ı���ȥ��β/��Ԫ�������/ȫ�ǿո����ǡ�Trim��
Private Function TrimVisible(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    TrimVisible = Trim$(s)
End Function

