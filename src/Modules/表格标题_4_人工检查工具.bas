Attribute VB_Name = "������_4_�˹���鹤��"
Sub ȫ�ı���ʽ������()

    Set doc = ActiveDocument
    Set tb = doc.Tables
    
    i = tb.Count
    
    For r = 1 To i
    tb(r).Select
    Next
    
End Sub
    
    
    
    
