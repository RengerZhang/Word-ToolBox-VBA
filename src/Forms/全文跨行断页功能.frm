VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ȫ�Ŀ��ж�ҳ���� 
   Caption         =   "UserForm1"
   ClientHeight    =   840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3200
   OleObjectBlob   =   "ȫ�Ŀ��ж�ҳ����.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "ȫ�Ŀ��ж�ҳ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
    ' ���� CheckBox ��״̬�����ÿ��ж�ҳ����
    If CheckBox1.Value = True Then
        ' ���ÿ��ж�ҳ����
        SetTableRowPageBreak True
    Else
        ' ���ÿ��ж�ҳ����
        SetTableRowPageBreak False
    End If
End Sub

Sub SetTableRowPageBreak(enable As Boolean)
    Dim tbl As Table
    Dim row As row
    
    ' �����ĵ��е����б��
    For Each tbl In ActiveDocument.Tables
        ' ����ÿ��������
        For Each row In tbl.rows
            ' ��ȡÿһ�еĵ�һ����Ԫ��
            Set firstCell = row.Cells(1)
            
            ' ���û�ȡ���еĿ��ж�ҳ����
            firstCell.row.AllowBreakAcrossPages = enable
        Next row
    Next tbl
End Sub

