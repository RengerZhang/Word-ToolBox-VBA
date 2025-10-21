Attribute VB_Name = "纵向合并单元格拆分"
Sub Test拆分选中的表格的合并单元格()
    Dim tbl As Table
    Dim i, j, errorCount As Integer
    Dim 开始错误的行号 As Integer
    Dim zonghang As Integer, zonglie As Integer
    Dim currentCell As cell
    Dim mergedRows As Integer
    
    ' 获取选中的表格
    If Selection.Tables.Count = 0 Then
        MsgBox "请先选择一个表格！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Set tbl = Selection.Tables(1) ' 获取选中的第一个表格
    
    zonghang = tbl.rows.Count ' 总行数
    zonglie = tbl.Columns.Count ' 总列数
    
    
    ' 遍历每一列
    For j = 1 To zonglie
        On Error Resume Next ' 启用错误处理
        
        ' 遍历每一行
        For i = 1 To zonghang
            Set currentCell = tbl.cell(i, j)
            
            ' 判断是否为合并单元格
            currentCell.Select
            If Err.Number <> 0 Then
                ' 如果遇到合并单元格，记录合并单元格的起始位置
                If errorCount = 0 Then
                    开始错误的行号 = i - 1
                End If
                
                ' 增加错误计数
                errorCount = errorCount + 1
                
                ' 给合并单元格标记颜色
                currentCell.Shading.BackgroundPatternColor = wdColorYellow ' 设置为黄色
                
                ' 清除错误
                Err.Clear
            Else
                ' 如果当前是非合并单元格，处理合并单元格的结束
                If errorCount > 0 Then
                    ' 输出合并单元格的位置和合并的行数
                    合并行数 = errorCount + 1
                    Debug.Print "合并的单元格在：第" & 开始错误的行号 & "行，第" & j & "列，合并了" & 合并行数 & "行"
                    
                    ' 记录合并单元格区域的行数
                    mergedRows = 合并行数
                    
                    ' 拆分合并单元格
                    Call 拆分单元格(CInt(开始错误的行号), CInt(mergedRows), CInt(j)) ' 强制转换为整数类型
                    
                    ' 重置错误计数器
                    errorCount = 0
                    开始错误的行号 = 0
                End If
            End If
        Next i
        
        On Error GoTo 0 ' 禁用错误处理
    Next j
    
     ' 循环结束后，选择第一行并将其文字加粗
    tbl.rows(1).Range.Font.bold = True
    Debug.Print "第一行文字已加粗"
    
End Sub


' 拆分合并的单元格
Sub 拆分单元格(startRow As Integer, mergedRows As Integer, col As Integer)
    Dim currentCell As cell
    
    ' 选中当前合并的单元格
    Set currentCell = Selection.Tables(1).cell(startRow, col)
    
    ' 输出拆分的单元格位置
    Debug.Print "正在拆分：第" & startRow & "行第" & col & "列的单元格"
    
    ' 选中合并单元格
    currentCell.Select
    
    ' 执行拆分操作，拆分合并单元格
    Selection.Cells.Split NumRows:=mergedRows, NumColumns:=1, MergeBeforeSplit:=False
    
    Debug.Print "合并单元格已拆分，第" & startRow & "行第" & col & "列的单元格"
End Sub

