VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 全文跨行断页功能 
   Caption         =   "UserForm1"
   ClientHeight    =   840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3200
   OleObjectBlob   =   "全文跨行断页功能.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "全文跨行断页功能"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
    ' 根据 CheckBox 的状态来设置跨行断页功能
    If CheckBox1.Value = True Then
        ' 启用跨行断页功能
        SetTableRowPageBreak True
    Else
        ' 禁用跨行断页功能
        SetTableRowPageBreak False
    End If
End Sub

Sub SetTableRowPageBreak(enable As Boolean)
    Dim tbl As Table
    Dim row As row
    
    ' 遍历文档中的所有表格
    For Each tbl In ActiveDocument.Tables
        ' 遍历每个表格的行
        For Each row In tbl.rows
            ' 获取每一行的第一个单元格
            Set firstCell = row.Cells(1)
            
            ' 设置或取消行的跨行断页功能
            firstCell.row.AllowBreakAcrossPages = enable
        Next row
    Next tbl
End Sub

