Attribute VB_Name = "表格标题_1_统一表标题样式"
Sub 表标题样式统一()
    Dim doc As Document
    Dim Para As Paragraph
    Dim lvl As Long
    Dim text As String
    Dim tbl As Table
    Dim rng As Range
    
    Set doc = ActiveDocument

      '======== 表格标题处理（容错：允许多个空白段）========
    For Each tbl In doc.Tables
        Set rng = tbl.Range
        rng.Collapse wdCollapseStart
        
        Dim prevPara As Paragraph
        Set prevPara = rng.Paragraphs(1).Previous
        
        Do While Not prevPara Is Nothing
            Dim t As String
            t = 清理段落文本(prevPara.Range.text)
            If Len(t) > 0 Then
                prevPara.Style = doc.Styles("表格标题")
                Exit Do
            End If
            Set prevPara = prevPara.Previous
        Loop
    Next tbl
    
End Sub
'——— 工具：清理段落可见文本（去掉段尾/单元格标记、全角空格并 Trim）
Private Function 清理段落文本(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")                ' 单元格结束符
    s = Replace(s, ChrW(&H3000), " ")         ' 全角空格转半角
    清理段落文本 = Trim$(s)
End Function

