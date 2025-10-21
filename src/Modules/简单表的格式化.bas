Attribute VB_Name = "简单表的格式化"
Sub FormatTable()
    ' 声明变量
    Dim doc As Document
    Dim tbl As Table
    Dim titleRow As row
    Dim i As Integer, j As Integer ' 循环行和列
    Dim cell As cell ' 单个单元格
    Dim minRowHeight As Single ' 行高最小值（0.6cm）
    
    ' 1. 检查是否选中表格
    On Error Resume Next
    Set tbl = Selection.Tables(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "请先选中一个表格再运行此宏！", vbExclamation, "提示"
        Exit Sub
    End If
    Set doc = ActiveDocument
    
    ' 计算0.6厘米对应的磅值（1厘米≈28.35磅）
    minRowHeight = CentimetersToPoints(0.6)
    
    ' 2. 清除表格内所有格式（包括底色）
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            Set cell = tbl.cell(i, j)
            With cell.Range
                ' 清除文字属性
                With .Font
                    .NameFarEast = ""
                    .NameAscii = ""
                    .bold = False
                    .Italic = False
                    .Underline = wdUnderlineNone
                    .Color = wdColorAutomatic
                    .Size = 10
                End With
                
                ' 清除段落属性
                With .ParagraphFormat
                    .alignment = wdAlignParagraphLeft
                    .LeftIndent = 0
                    .RightIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceSingle
                End With
                
                ' 清除单元格底色
                .Shading.BackgroundPatternColor = wdColorAutomatic
            End With
        Next j
    Next i
    
    ' 3. 表格整体设置
    tbl.AutoFitBehavior wdAutoFitWindow ' 适应窗口宽度
    tbl.AllowPageBreaks = True ' 允许跨页
    
    ' 4. 设置所有单元格格式（核心修改：边距和居中）
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            Set cell = tbl.cell(i, j)
            
            ' 设置单元格边距（上下边距为0）
            With cell
                .TopPadding = 0 ' 上边距0磅
                .BottomPadding = 0 ' 下边距0磅
                ' 左右边距保持默认（如需调整可添加：.LeftPadding = 0 或 .RightPadding = 0）
            End With
            
            ' 设置文字格式
            With cell.Range.Font
                .NameFarEast = "宋体" ' 中文宋体
                .NameAscii = "Times New Roman" ' 西文Times New Roman
                .Size = 10 ' 统一字号
            End With
            
            ' 文字水平居中（左右居中）
            cell.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
            
            ' 文字垂直居中（高度居中）
            cell.VerticalAlignment = wdCellAlignVerticalCenter
        Next j
    Next i
    
    ' 5. 设置标题行（第一行）
    Set titleRow = tbl.rows(1)
    titleRow.HeadingFormat = True ' 跨页重复标题行
    titleRow.Range.Font.bold = True ' 标题行加粗
    
    ' 6. 设置行高规则（核心修改：最小值0.6cm）
    For i = 1 To tbl.rows.Count
        With tbl.rows(i)
            .HeightRule = wdRowHeightAtLeast ' 行高至少为指定值（内容多时自动扩展）
            .Height = minRowHeight ' 最小值0.6cm
        End With
    Next i
    
    ' 7. 设置边框（外框1.5磅，内框0.5磅）
    With tbl.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .OutsideLineWidth = wdLineWidth150pt ' 1.5磅
        .OutsideColor = wdColorBlack
        
        .InsideLineStyle = wdLineStyleSingle
        .InsideLineWidth = wdLineWidth050pt ' 0.5磅
        .InsideColor = wdColorBlack
    End With
    
    ' 完成提示
    MsgBox "表格格式设置完成！" & vbCrLf & _
           "1. 行高最小值0.6cm（内容多时自动扩展）" & vbCrLf & _
           "2. 单元格上下边距为0，文字水平+垂直居中" & vbCrLf & _
           "3. 外框1.5磅，内框0.5磅，中文宋体，西文Times New Roman" & vbCrLf & _
           "4. 标题行加粗并跨页重复", vbInformation, "完成"
End Sub

