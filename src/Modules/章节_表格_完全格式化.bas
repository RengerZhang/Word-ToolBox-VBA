Attribute VB_Name = "章节_表格_完全格式化"
Sub 章节表格格式化工具()
    Dim oRow As row
    Dim oCell As cell
    Dim tb As Table
    Dim myParagraph As Paragraph, n As Integer
    Dim progressForm As progressForm ' 引用之前创建的 ProgressForm
    Dim 当前章节 As Range
    Dim 表格计数 As Integer
    Dim 当前节 As Integer
    
    ' （一）调用“标准化表格样式”函数
    EnsureStandardTableStyle
    
    
    ' (二)获取章节表格数量和表格
    ' 获取光标所在的节
    当前节 = Selection.Sections(1).Index
    
    ' 获取当前节的范围
    Set 当前章节 = ActiveDocument.Sections(当前节).Range
    
    ' 计算该节中的表格数量
    表格计数 = 当前章节.Tables.Count
    i = 表格计数
    
    ' 创建并显示进度窗体
    Set progressForm = New progressForm
    progressForm.Show vbModeless ' 非模态窗体，这样代码可以继续执行
    progressForm.UpdateProgressBar 0, "Processing table 1 of " & i   ' 初始化进度窗体
    progressForm.TextBoxStatus.text = "本节共有 " & i & " 个表格，现在开始格式化..."

    
    For r = 1 To i
        If progressForm.stopFlag Then
            ' 如果点击了强制停止按钮，退出循环
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "操作已停止，正在退出..."
            Exit For
        End If
        
        tbl = 当前章节.Tables(r)
        Set tb = 当前章节.Tables(r)
        
        tbl.Select
        Selection.Style = "标准化表格样式"
        '    调用表格属性设置函数
        Call 表格属性设置(ActiveDocument.Tables(r))
        
        
        
        '  遍历单元格设置内框线,为了防止合并单元格现象
        For Each oCell In tbl.Cells
            oCell.Select
            With Selection
                .Borders.OutsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineWidth = wdLineWidth050pt
            End With
            
            Selection.SelectRow
            Selection.rows.AllowBreakAcrossPages = enable
            
            n = 1
            For Each myParagraph In Selection.Paragraphs
                If Len(Trim(myParagraph.Range)) = 1 Then
                    myParagraph.Range.Delete
                    n = n + 1
                End If
            Next
            
        Next oCell

        '  选中表格，设置外框线
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth150pt ' 外框1.5磅
            .OutsideColor = wdColorBlack
        End With
        
        '  设置首行加粗并且每页重复
        tbl.Select
        Selection.rows.HeadingFormat = False
        tbl.Cells(1).Select
        Selection.SelectRow
        Selection.Range.bold = True
        Selection.rows.HeadingFormat = True
        
       
        ' 更新进度条和状态文本框
        progressForm.UpdateProgressBar (r / i) * 200, "Processing table " & r & " of " & i
        
        ' 确保窗体更新
        DoEvents
        
    Next r
    ' 完成后，在 TextBox 中显示完成消息
    
    ' 追加记录到 TextBox
    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "表格格式调整完毕！"
    'progressForm.Hide
    
    Exit Sub
    
    ' 详细操作提示信息
    MsgBox "表格格式调整完成，已执行以下操作：" & vbCrLf & _
        "1. 应用""表格""样式" & vbCrLf & _
        "2. 自动适应窗口宽度" & vbCrLf & _
        "3. 设置单元格上下边距为0" & vbCrLf & _
        "4. 行居中对齐，取消文字环绕" & vbCrLf & _
        "5. 字体设置：宋体(中文)、Times New Roman(英文)，10.5号字" & vbCrLf & _
        "6. 单元格内容垂直和水平居中" & vbCrLf & _
        "7. 段落设置：单倍行距，无前后间距，无缩进" & vbCrLf & _
        "8. 清除底色，行高设为最小值0.6cm" & vbCrLf & _
        "9. 设置内框线(0.5磅)和外框线(1.5磅)" & vbCrLf & _
        "10. 禁止表格行跨页断行" & vbCrLf & _
        "11. 删除单元格内空段落" & vbCrLf & _
        "12. 首行加粗并设置为每页重复标题行", _
        vbInformation, "操作完成"
    
End Sub
