Attribute VB_Name = "全文_表格_统一字体样式"
Sub 全文表格样式格式化()
    Dim tb As Table
    Dim myParagraph As Paragraph, N As Integer
    Dim progressForm As progressForm ' 引用之前创建的 ProgressForm
    
    ' 调用“标准化表格样式”函数
    EnsureStandardTableStyle
    
    ' 计算全文表格数量
    i = ActiveDocument.Tables.Count
    
    ' 创建并显示进度窗体
    Set progressForm = New progressForm
    progressForm.Show vbModeless ' 非模态窗体，这样代码可以继续执行
    progressForm.TextBoxStatus.text = "全文共有" & i & "个表格，现在开始格式化表格样式..."
     progressForm.UpdateProgressBar 0, "Processing table 1 of " & i   ' 初始化进度窗体
    
    
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "操作已停止，正在退出..."
            Exit For
        End If
        
        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)
        
        tbl.Select
        Selection.Style = "标准化表格样式"
       
        ' 更新进度条和状态文本框
        progressForm.UpdateProgressBar (r / i) * 200, "Processing table " & r & " of " & i
        ' 确保窗体更新
        DoEvents
        
    Next r

    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "表格样式设置完毕！"
    
    Exit Sub
    
End Sub
