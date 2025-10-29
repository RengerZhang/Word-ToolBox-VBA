Attribute VB_Name = "中交标准化表格格式化工具"
Sub AdjustTableFormat()
    Dim oRow As row
    Dim oCell As cell
    Dim tb As Table
    Dim myParagraph As Paragraph, N As Integer
    
    ' 调用“标准化表格样式”函数
    EnsureStandardTableStyle
    
    Selection.Tables(1).Select
    Selection.Style = "标准化表格样式"
    
    
    tbl = Selection.Tables(1)
    Set tb = Selection.Tables(1)
    
    ' 检查是否选中了表格
    On Error Resume Next
    Set tbls = Selection.Tables  ' 获取选中内容中的表格集合
    On Error GoTo 0
    
    ' 如果没有选中表格，提示用户并退出
    If tbls.Count = 0 Then
        MsgBox "请先选中一个表格再运行本宏！", vbExclamation, "无选中表格"
        Exit Sub
    End If
    
    
    '    Selection.ParagraphFormat.Reset
    '    tb.AutoFitBehavior (wdAutoFitContent)
    tb.AutoFitBehavior (wdAutoFitWindow)
    tb.TopPadding = PixelsToPoints(0, True) '设置上边距为0
    tb.BottomPadding = PixelsToPoints(0, True) '设置下边距为0
    tb.LeftPadding = PixelsToPoints(0, True) '设置上边距为0
    tb.RightPadding = PixelsToPoints(0, True) '设置下边距为0
    
    '格式化表格
    tbl.Select
    With Selection
        .rows.alignment = wdAlignRowCenter
        .rows.WrapAroundText = False
        .Font.NameFarEast = ""
        .Font.NameAscii = ""
        .Range.bold = False
        .Range.Font.NameAscii = "Times New Roman"
        .Range.Font.NameFarEast = "宋体"
        .Range.Font.Size = 10.5
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter ' 单元格垂直居中
        .ParagraphFormat.CharacterUnitFirstLineIndent = 0
        .ParagraphFormat.alignment = wdAlignParagraphCenter ' 居中对齐
        .ParagraphFormat.SpaceBefore = 0 ' 段前
        .ParagraphFormat.SpaceAfter = 0 ' 段后
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle ' 单倍行距
        .ParagraphFormat.CharacterUnitFirstLineIndent = 0 ' 左侧缩进
        .ParagraphFormat.LeftIndent = 0 ' 左侧缩进
        .ParagraphFormat.RightIndent = 0 ' 右侧缩进
        .Shading.BackgroundPatternColor = wdColorAutomatic ' 清除底色
        .rows.HeightRule = wdRowHeightAtLeast
        .rows.Height = CentimetersToPoints(0.6)
    End With
    
    '  遍历单元格设置内框线,为了防止合并单元格现象
    For Each oCell In tbl.Cells
        oCell.Select
        With Selection
            .Borders.OutsideLineStyle = wdLineStyleSingle
            .Borders.OutsideLineWidth = wdLineWidth050pt
        End With
        
        Selection.SelectRow
        Selection.rows.AllowBreakAcrossPages = enable
        
        N = 1
        For Each myParagraph In Selection.Paragraphs
            If Len(Trim(myParagraph.Range)) = 1 Then
                myParagraph.Range.Delete
                N = N + 1
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
