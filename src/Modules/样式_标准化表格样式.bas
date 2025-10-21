Attribute VB_Name = "样式_标准化表格样式"
Option Explicit

'========================================
' 预处理：按“文字单元格 = 总单元格 ? 含图单元格”规则打标
' 判定：
'   n = 图片总数（Inline + Shape）
'   若 估算文字单元格 ≤ (n + 1) → 【图片定位表】；否则 → 【标准表格样式】
'   无图（n=0）→ 直接【标准表格样式】
' 处理后在“立即窗口”(Ctrl+G)输出每表明细
'========================================
Public Sub 预处理_标记图片表与普通表()
    '（一）样式名
    Const S_TABLE_PIC As String = "图片定位表"
    Const S_TABLE_NOR As String = "标准表格样式"

    '（二）保证样式存在（不改外观）
    Dim doc As Document: Set doc = ActiveDocument
    EnsureTableStyleOnly doc, S_TABLE_PIC
    EnsureTableStyleOnly doc, S_TABLE_NOR

    '（三）逐表判定
    Dim i As Long, tb As Table
    Dim nInline As Long, nShape As Long, nImgObj As Long
    Dim totalCells As Long, imgCellCnt As Long, txtCellEst As Long
    Dim threshold As Long, applied As String, imgCoords As String

    For i = 1 To doc.Tables.Count
        Set tb = doc.Tables(i)

        ' 1) 图片对象总数（n）
        nInline = tb.Range.InlineShapes.Count
        nShape = SafeShapeCount_InRange(tb.Range)
        nImgObj = nInline + nShape

        ' 2) 单元格层统计：总单元格数 & 含图单元格数
        totalCells = tb.Range.Cells.Count                           ' 兼容合并单元格
        imgCellCnt = CountImageCells(tb, imgCoords)                 ' 逐单元格检查是否含图
        txtCellEst = totalCells - imgCellCnt                        ' 估算文字单元格数

        ' 3) 判定：阈值 = n + 1
        threshold = nImgObj + 1
        If (nImgObj > 0) And (txtCellEst <= threshold) Then
            tb.Style = S_TABLE_PIC
            applied = S_TABLE_PIC
        Else
            tb.Style = S_TABLE_NOR
            applied = S_TABLE_NOR
        End If

        ' 4) 调试输出
        Debug.Print "表#" & i & _
                    " | 尺寸=" & tb.rows.Count & "x" & tb.Columns.Count & _
                    " | 总单元格=" & totalCells & _
                    " | 图片对象 n=Inline:" & nInline & "+Shape:" & nShape & "=" & nImgObj & _
                    " | 含图单元格=" & imgCellCnt & _
                    " | 估算文字单元格=" & txtCellEst & " ≤? 阈值(n+1)=" & threshold & _
                    " | 判定=" & applied & _
                    IIf(Len(imgCoords) > 0, " | 含图坐标:" & imgCoords, "")
    Next i

    MsgBox "预处理完成（按“总单元格?含图单元格”计算文字单元格）。详情见立即窗口。", vbInformation
End Sub

'========================================
' 工具：统计 Range 内“浮动形状”个数（无形状时不报错）
'========================================
Private Function SafeShapeCount_InRange(ByVal rng As Range) As Long
    On Error Resume Next
    SafeShapeCount_InRange = rng.ShapeRange.Count
    On Error GoTo 0
End Function

'========================================
' 工具：统计“含图单元格”数量，并返回坐标清单 "(r,c),(r,c)..."
' 规则：单元格内 InlineShapes.Count + ShapeRange.Count > 0 即视为“含图”
'========================================
Private Function CountImageCells(ByVal tb As Table, ByRef coords As String) As Long
    Dim c As cell, n As Long, buf As String
    For Each c In tb.Range.Cells
        If (c.Range.InlineShapes.Count > 0) Or (SafeShapeCount_InRange(c.Range) > 0) Then
            n = n + 1
            If Len(buf) > 0 Then buf = buf & ","
            ' Word 的 Cell 对象支持 RowIndex / ColumnIndex
            buf = buf & "(" & c.rowIndex & "," & c.ColumnIndex & ")"
        End If
    Next c
    coords = buf
    CountImageCells = n
End Function

'========================================
' 兜底：保证样式存在且为“表格样式”（不改外观）
'========================================
Private Sub EnsureTableStyleOnly(ByVal doc As Document, ByVal styleName As String)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    If Not st Is Nothing Then
        If st.Type <> wdStyleTypeTable Then
            st.Delete
            Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
        End If
    Else
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
    End If
End Sub


