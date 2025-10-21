Attribute VB_Name = "计算表格数量"
Sub 计算表格数量()
    Dim 当前章节 As Range
    Dim 表格计数 As Integer
    Dim 当前节 As Integer

    ' 获取光标所在的节
    当前节 = Selection.Sections(1).Index
    
    ' 获取当前节的范围
    Set 当前章节 = ActiveDocument.Sections(当前节).Range
    
    ' 计算该节中的表格数量
    表格计数 = 当前章节.Tables.Count
    ' 输出表格数量
    MsgBox "当前章节的表格数量是: " & 表格计数
End Sub
