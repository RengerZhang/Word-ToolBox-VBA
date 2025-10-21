Attribute VB_Name = "中交标准化表格设置2"
Sub FormatTable()
    ' 声明变量
    Dim tbl As Table
    Dim titleRow As row
    Dim i As Integer, j As Integer, k As Integer
    Dim minRowHeight As Single
    Dim currentCell As cell
    Dim targetRow As row
    Dim processedRows As New Collection ' 用于存储已处理的行，避免重复设置
    Dim originalRange As Range ' 用于保存原始选择范围
    Dim tableRange As Range ' 表格整体范围
    Dim firstRowRange As Range ' 第一行范围
    
    ' 保存原始选择范围，避免处理过程中改变用户选择
    Set originalRange = Selection.Range
    
    ' 1. 检查是否选中表格
    On Error Resume Next
    Set tbl = Selection.Tables(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "请先选中一个表格再运行此宏！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 计算0.6厘米对应的磅值（1厘米≈28.35磅）
    minRowHeight = CentimetersToPoints(0.6)
    
    ' 2. 清除表格内所有格式（包括底色）
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            ' 处理合并单元格可能导致的单元格不存在问题
            On Error Resume Next
            Set currentCell = tbl.cell(i, j)
            On Error GoTo 0
            
            If Not currentCell Is Nothing Then
                With currentCell.Range
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
                
                ' 释放对象
                Set currentCell = Nothing
            End If
        Next j
    Next i
    
    ' 3. 表格整体设置
    tbl.AutoFitBehavior wdAutoFitWindow ' 适应窗口宽度
    tbl.AllowPageBreaks = True ' 允许跨页
    
    ' 4. 设置所有单元格格式（边距、文字格式、居中）
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            On Error Resume Next
            Set currentCell = tbl.cell(i, j)
            On Error GoTo 0
            
            If Not currentCell Is Nothing Then
                ' 设置单元格边距（上下边距为0）
                With currentCell
                    .TopPadding = 0 ' 上边距0磅
                    .BottomPadding = 0 ' 下边距0磅
                End With
                
                ' 设置文字格式
                With currentCell.Range.Font
                    .NameFarEast = "宋体" ' 中文宋体
                    .NameAscii = "Times New Roman" ' 西文Times New Roman
                    .Size = 10 ' 统一字号
                End With
                
                ' 文字水平居中
                currentCell.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
                
                ' 文字垂直居中
                currentCell.VerticalAlignment = wdCellAlignVerticalCenter
                
                ' 释放对象
                Set currentCell = Nothing
            End If
        Next j
    Next i
    
    ' 2. 通过表格范围获取第一行，避免直接使用 Rows(1)
    On Error Resume Next
    ' 获取表格整体范围
    Set tableRange = tbl.Range
    ' 从表格范围中截取第一行的范围（关键改进）
    Set firstRowRange = tableRange.rows(1).Range
    On Error GoTo 0
    
    If Not firstRowRange Is Nothing Then
        ' 3. 从第一行范围中提取行对象
        On Error Resume Next
        Set titleRow = firstRowRange.rows(1)
        On Error GoTo 0
'
        ' 4. 验证行对象并设置标题行属性
        If Not titleRow Is Nothing Then
            titleRow.HeadingFormat = True ' 跨页重复
            titleRow.Range.Font.bold = True ' 加粗
        Else
            ' 降级方案：直接操作第一行范围的格式
            firstRowRange.Font.bold = True
            ' 跨页重复属性需要行对象，此处提示可能失效
            Debug.Print "警告：无法设置跨页重复，仅完成加粗"
        End If
    Else
        ' 终极降级：直接通过单元格范围设置格式
        On Error Resume Next
        ' 操作表格第一行第一个单元格所在的行范围
        tbl.cell(1, 1).Range.rows(1).Font.bold = True
        tbl.cell(1, 1).Range.rows(1).HeadingFormat = True
        On Error GoTo 0
        Debug.Print "使用单元格范围降级设置标题行"
    End If
    

    
    ' 恢复原始选择范围
    originalRange.Select
    
    ' 6. 设置行高规则（处理合并单元格的行访问问题）
    ' 收集所有唯一行，避免重复设置
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            On Error Resume Next
            Set currentCell = tbl.cell(i, j)
            On Error GoTo 0
            
            If Not currentCell Is Nothing Then
                ' 可靠的行获取方式
                Dim rowIndex As Integer
                rowIndex = currentCell.rowIndex
                On Error Resume Next
                Set targetRow = tbl.rows(rowIndex)
                On Error GoTo 0
                
                If Not targetRow Is Nothing Then
                    ' 检查行是否已处理
                    Dim isExists As Boolean
                    isExists = False
                    For Each existingRow In processedRows
                        If existingRow = targetRow.Index Then
                            isExists = True
                            Exit For
                        End If
                    Next
                    If Not isExists Then
                        On Error Resume Next
                        processedRows.Add targetRow.Index, CStr(targetRow.Index)
                        ' 设置行高
                        targetRow.HeightRule = wdRowHeightAtLeast
                        targetRow.Height = minRowHeight
                        On Error GoTo 0
                    End If
                Else
                    ' 记录错误信息
                    Debug.Print "获取行对象失败，行索引：" & rowIndex & "，单元格位置：第" & i & "行第" & j & "列"
                End If
                
                ' 释放对象
                Set currentCell = Nothing
                Set targetRow = Nothing
            End If
        Next j
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
           "4. 标题行（含合并行）加粗并跨页重复", vbInformation, "完成"
End Sub


