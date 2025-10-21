Attribute VB_Name = "全文_表格_标题加粗"
Sub 全文表格首行加粗()
    Dim tb As Table
    Dim myParagraph As Paragraph, n As Integer
    Dim progressForm As progressForm ' 引用之前创建的 ProgressForm
    
    
    ' 计算全文表格数量
    i = ActiveDocument.Tables.Count
    
    ' 创建并显示进度窗体
    Set progressForm = New progressForm
    progressForm.Show vbModeless ' 非模态窗体，这样代码可以继续执行
    progressForm.TextBoxStatus.text = "全文共有" & i & "个表格，现在开始加粗标题..."
     progressForm.UpdateProgressBar 0, "Processing table 1 of " & i   ' 初始化进度窗体
    
    
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "操作已停止，正在退出..."
            Exit For
        End If
        
        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)

        '  设置首行加粗并且每页重复
        tbl.rows.Select
        Selection.rows.HeadingFormat = wdUndefined
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

    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "表格标题加粗完毕！"
    
    Exit Sub
    
End Sub

'==========================================================
' 当前表格格式控制（光标所在的那一张表）
' 参数：
'   thickOuter   外框加粗（True=1.5磅；False=0.5磅）
'   firstRowBold 首行加粗
'   headerRepeat 首行每页重复
'   allowBreak   允许跨页断行（整表）
'   fontSizePt   整表字号（磅）
'==========================================================
Public Sub 当前表格格式设置(thickOuter As Boolean, _
                           firstRowBold As Boolean, _
                           headerRepeat As Boolean, _
                           allowBreak As Boolean, _
                           fontSizePt As Single)
    Dim tb As Table
    '（一）拿到光标所在表格
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "光标未在表格中。", vbExclamation
        Exit Sub
    End If
    Set tb = Selection.Tables(1)

    '（二）整表字号（按你“只控一个值”的思路）
    tb.Range.Font.Size = fontSizePt

    '（三）内框线：固定 0.5 磅（一行写死）
    tb.Borders.InsideLineStyle = wdLineStyleSingle
    tb.Borders.InsideLineWidth = wdLineWidth050pt

    '（四）外框线：开=1.5 磅；关=0.5 磅（直接控这里）
    With tb.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .OutsideLineWidth = IIf(thickOuter, wdLineWidth150pt, wdLineWidth050pt)
        .OutsideColor = wdColorBlack
    End With

    '（五）首行加粗 & 首行重复（直接控这两处）
    If tb.rows.Count > 0 Then
        tb.rows(1).Range.bold = firstRowBold
        tb.rows(1).HeadingFormat = headerRepeat
    End If

    '（六）整表是否允许跨页断行
    tb.rows.AllowBreakAcrossPages = allowBreak
End Sub


