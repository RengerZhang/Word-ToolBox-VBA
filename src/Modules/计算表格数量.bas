Attribute VB_Name = "����������"
Sub ����������()
    Dim ��ǰ�½� As Range
    Dim ������ As Integer
    Dim ��ǰ�� As Integer

    ' ��ȡ������ڵĽ�
    ��ǰ�� = Selection.Sections(1).Index
    
    ' ��ȡ��ǰ�ڵķ�Χ
    Set ��ǰ�½� = ActiveDocument.Sections(��ǰ��).Range
    
    ' ����ý��еı������
    ������ = ��ǰ�½�.Tables.Count
    ' ����������
    MsgBox "��ǰ�½ڵı��������: " & ������
End Sub
