Attribute VB_Name = "全文_表格_跨行断页"
Sub SetTableRowPageBreak()
    Dim row As row
    
    Selection.Tables(1).Select
    tbl = Selection.Tables(1)
    Set tb = Selection.Tables(1)
    
    ' 检查是否选中了表格
    On Error Resume Next
    Set tbls = Selection.Tables  ' 获取选中内容中的表格集合
    On Error GoTo 0
    
    ' 如果没有选中表格，提示用户并退出
    If tbls.Count = 0 Then
        MsgBox "请先选中一个表格再运行本宏！", vbExclamation, "无选中表格"
        Exit Sub
    End If
    
    
    ' 遍历每个表格的行
    For Each oCell In tbl.Cells
        oCell.Select
        Selection.SelectRow
        Selection.rows.AllowBreakAcrossPages = enable
    Next
    
End Sub

