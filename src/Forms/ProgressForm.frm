VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "表格格式化处理进度"
   ClientHeight    =   3500
   ClientLeft      =   90
   ClientTop       =   410
   ClientWidth     =   4990
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProgressForm 的代码
Public Event OnClose()  ' 定义关闭事件（供外部订阅）
Public stopFlag As Boolean ' 用于控制强制停止
'
'
'Private Sub TextBoxStatus_Change()
'
'End Sub

Private Sub TextBoxStatus_Change()

End Sub

Private Sub UserForm_Initialize()
    ' 初始化控件
    Me.LabelProgressBar.width = 200
    Me.LabelProgressBar.Height = 24
    Me.LabelProgressBar.Top = 5
    Me.LabelProgressBar.Left = 5
    Me.FrameProgress.width = 0 ' 初始进度为 0
    Me.FrameProgress.Top = 5.5
    Me.FrameProgress.Left = 5.5
    Me.FrameProgress.Height = 25
    Me.TextBoxStatus.width = 235
    Me.TextBoxStatus.Height = 100
    Me.TextBoxStatus.MultiLine = True ' 支持多行
    Me.TextBoxStatus.ScrollBars = fmScrollBarsVertical ' 启用垂直滚动条
    Me.TextBoxStatus.Locked = True ' 使 TextBox 为只读模式
    
    ' 第一行显示总表格数量及开始格式化提示
'    Me.TextBoxStatus.Text = "该文档总共有 " & ActiveDocument.Tables.Count & " 个表格，现在开始格式化..."
    
    ' 初始化进度条
    Me.FrameProgress.width = 0
    Me.LabelPercentage.caption = "0%" ' 初始化百分比显示
End Sub
' 供“页面工具”调用的初始化入口
Public Sub InitForPageSetting(ByVal totalSections As Long, ByVal startSectionIndex As Long)
    Me.caption = "全文页面设置处理进度"
    Me.FrameProgress.width = 0
    Me.LabelPercentage.caption = "0%"
    Me.TextBoxStatus.text = _
        "本文共有 " & totalSections & " 节，正文从第 " & startSectionIndex & " 节开始。"
End Sub

' 更新进度条和状态文本框
Public Sub UpdateProgressBar(ByVal progress As Integer, ByVal statusMessage As String)
    If stopFlag Then
'        Me.TextBoxStatus.Text = "操作被强制停止！"
        progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "操作被强制停止！"
        Exit Sub
    End If
    
    ' 更新进度条
    Me.FrameProgress.width = progress
    ' 更新百分比
    Me.LabelPercentage.caption = progress / 2 & "%" ' 200px 为满进度，进度条最大宽度

    ' 追加记录到 TextBox
    Me.TextBoxStatus.text = Me.TextBoxStatus.text & vbCrLf & statusMessage
    
   ' ==== 强制滚动到最后一行（不会抢走按钮焦点）====
    Dim ctrl As MSForms.Control
    Set ctrl = Me.ActiveControl
    
    Me.TextBoxStatus.SetFocus
    Me.TextBoxStatus.SelStart = Len(Me.TextBoxStatus.text)
    Me.TextBoxStatus.SelLength = 0
    
    ' 还原原来焦点
    If Not ctrl Is Nothing Then
        On Error Resume Next
        ctrl.SetFocus
        On Error GoTo 0
    End If
    ' ================================================

    
    DoEvents ' 确保更新
End Sub

' 强制停止按钮点击事件
Private Sub btnStop_Click()
    stopFlag = True
'    Me.TextBoxStatus.Text = "操作已被强制停止！"
    Me.TextBoxStatus.text = Me.TextBoxStatus.text & vbCrLf & "操作已被强制停止！"
    Me.LabelPercentage.caption = "终止！"
    RaiseEvent OnClose    ' 触发关闭事件
End Sub

' 关闭按钮点击事件
Private Sub btnClose_Click()
    Me.Hide ' 隐藏窗体
    RaiseEvent OnClose    ' 触发关闭事件
End Sub
Private Sub ScrollTextBoxToEnd(tb As MSForms.TextBox)
    Dim ctrl As Control
    ' 先记住当前焦点控件
    Set ctrl = Me.ActiveControl
    
    tb.SetFocus
    tb.SelStart = Len(tb.text)
    tb.SelLength = 0
    
    ' 再把焦点还原回去
    If Not ctrl Is Nothing Then ctrl.SetFocus
End Sub
