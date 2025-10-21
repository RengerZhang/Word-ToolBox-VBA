Attribute VB_Name = "表格标题_4_人工检查工具"
Sub 全文表格格式化工具()

    Set doc = ActiveDocument
    Set tb = doc.Tables
    
    i = tb.Count
    
    For r = 1 To i
    tb(r).Select
    Next
    
End Sub
    
    
    
    
