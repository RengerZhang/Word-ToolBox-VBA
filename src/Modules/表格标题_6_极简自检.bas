Attribute VB_Name = "表格标题_6_极简自检"
Option Explicit

'——入口：极简自检（不改文档）
Public Sub 自检_表格标题_极简()
    Dim doc As Document: Set doc = ActiveDocument
    Dim a() As Long: a = BuildPrevStartArray(doc)

    With New 自检报告
        .LoadReportLite doc.Tables.Count, a   ' ← 只传表总数 + 每表前段Start
        .Show vbModeless
    End With
End Sub

'——为每个表计算“就近上一非空段”的 Start（找不到=-1）
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
        If (i And 31) = 0 Then DoEvents  ' 防卡顿
    Next
    BuildPrevStartArray = arr
End Function

'——就近上一非空段（允许多个空白段）
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

'——清理可见文本（去段尾/单元格结束符/全角空格→半角→Trim）
Private Function TrimVisible(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    TrimVisible = Trim$(s)
End Function

