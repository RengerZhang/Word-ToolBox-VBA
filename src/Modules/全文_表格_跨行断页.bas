Attribute VB_Name = "ȫ��_���_���ж�ҳ"
Sub SetTableRowPageBreak()
    Dim row As row
    
    Selection.Tables(1).Select
    tbl = Selection.Tables(1)
    Set tb = Selection.Tables(1)
    
    ' ����Ƿ�ѡ���˱��
    On Error Resume Next
    Set tbls = Selection.Tables  ' ��ȡѡ�������еı�񼯺�
    On Error GoTo 0
    
    ' ���û��ѡ�б����ʾ�û����˳�
    If tbls.Count = 0 Then
        MsgBox "����ѡ��һ����������б��꣡", vbExclamation, "��ѡ�б��"
        Exit Sub
    End If
    
    
    ' ����ÿ��������
    For Each oCell In tbl.Cells
        oCell.Select
        Selection.SelectRow
        Selection.rows.AllowBreakAcrossPages = enable
    Next
    
End Sub

