Attribute VB_Name = "全文_表格_完全格式化_跳过图片表"
' =======【新增】常量：样式名（放在模块顶部，过程外）=======
Const S_TABLE_PIC As String = "图片定位表"     ' 表格样式（图片表）
Const S_TABLE_NOR As String = "标准表格样式"   ' 表格样式（一般表）
Const S_PARA_IMG As String = "图片格式"        ' 段落样式：用于“含图单元格”
Const S_PARA_CAP As String = "图片标题"        ' 段落样式：用于“非图单元格”（标题）


'====================（A）新增：按参数直接执行 ====================
Public Sub 全文表格格式化_按参数1( _
    ByVal optThick As Boolean, _
    ByVal optHeadBold As Boolean, _
    ByVal optFontPt As Single, _
    Optional ByVal optFontName As String = "五号" _
)
    '（一）参数校验
    If optFontPt <= 0! Then
        MsgBox "字号参数无效。", vbExclamation
        Exit Sub
    End If
    '（二）进入核心
    全文表格格式化_核心 optThick, optHeadBold, optFontPt, optFontName
End Sub

' 只在“逐表处理”的地方加一个图片表分支；一般表原逻辑不动
Private Sub 全文表格格式化_核心( _
    ByVal optThick As Boolean, _
    ByVal optHeadBold As Boolean, _
    ByVal optFontPt As Single, _
    ByVal optFontName As String _
)
    Dim tb As Table
    Dim oCell As cell
    Dim myParagraph As Paragraph
    Dim N As Integer
    Dim progressForm As progressForm
    Dim i As Long, r As Long

    ' 样式准备（你原有的）
    EnsureStandardTableStyle
    ActiveDocument.Styles("标准化表格样式").Font.Size = optFontPt

    ' 进度窗体（你原有的）
    i = ActiveDocument.Tables.Count
    Set progressForm = New progressForm
    progressForm.Show vbModeless
    progressForm.TextBoxStatus.text = "全文共有 " & i & " 个表格，现在开始格式化..."
    progressForm.UpdateProgressBar 0, "Processing table 1 of " & i

    ' ===== 逐表处理 =====
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "操作已停止，正在退出..."
            Exit For
        End If

        Set tb = ActiveDocument.Tables(r)

        ' =========【新增】图片表判定（n图 → 阈值 n+1）=========
        Dim nInline As Long, nShape As Long, nImg As Long
        Dim totalCells As Long, imgCells As Long, txtEst As Long
        Dim isPic As Boolean

        nInline = tb.Range.InlineShapes.Count
        nShape = SafeShapeCount_InRange(tb.Range)   ' 浮动图个数（安全）
        nImg = nInline + nShape

        totalCells = tb.Range.Cells.Count           ' 兼容合并单元格
        imgCells = CountImageCells(tb)              ' 含图单元格数
        txtEst = totalCells - imgCells              ' 估算文字单元格
        isPic = (nImg > 0 And txtEst <= nImg + 1)

        If isPic Then
            ' =======【新增】图片表处理=======
            Call 格式化_图片定位表(tb)            ' 只做你要求的三件事
            ' 进度 & 下一张表
            progressForm.UpdateProgressBar CLng((r / i) * 200), _
                "Processing table " & r & " of " & i & "（图片定位表）"
            DoEvents
            GoTo NextTable_Continue                 ' 跳过一般表的原流程
        End If
        ' =========【新增】分支结束，下面进入“原有的一般表流程”=========

        ' ========【以下保持你的原代码不动】========
        ' 1) 应用样式 + 常规属性
        tb.Select: Selection.Style = "标准化表格样式"
        表格属性设置 tb                                   ' ← 你原来的过程名
        ' 2) 内框线固定 0.5 磅
        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)
        For Each oCell In tbl.Cells
            oCell.Select
            With Selection
                .Borders.OutsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineWidth = wdLineWidth050pt
            End With
            Selection.SelectRow
            Selection.rows.AllowBreakAcrossPages = False
            N = 1
            For Each myParagraph In Selection.Paragraphs
                If Len(Trim(myParagraph.Range)) = 1 Then
                    myParagraph.Range.Delete
                    N = N + 1
                End If
            Next
        Next oCell

        ' 3) 外框线：按开关 1.5 / 0.5
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = IIf(optThick, wdLineWidth150pt, wdLineWidth050pt)
            .OutsideColor = wdColorBlack
        End With

        ' 4) 首行加粗 + 每页重复
        tbl.Select
        Selection.rows.HeadingFormat = False
        tbl.Cells(1).Select
        Selection.SelectRow
        Selection.Range.bold = optHeadBold
        Selection.rows.HeadingFormat = True

        ' 5) 进度
        progressForm.UpdateProgressBar CLng((r / i) * 200), _
            "Processing table " & r & " of " & i
        DoEvents

NextTable_Continue:
    Next r

    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "表格格式调整完毕！"
End Sub


'====================（C）兼容保留：旧入口（弹窗取值→再调核心） ====================
Public Sub 全文表格格式化工具1()
    Dim dlg As 标准化格式工具箱
    Set dlg = New 标准化格式工具箱
    dlg.Show vbModeless
    If dlg.Canceled Then
        MsgBox "已取消。", vbInformation
        Exit Sub
    End If

    全文表格格式化_核心 _
        dlg.SelectedThickOuter, _
        dlg.SelectedFirstRowBold, _
        dlg.SelectedFontSizePt, _
        dlg.SelectedFontSizeName
End Sub

'——安全统计某范围内浮动图数（无图不报错）
Private Function SafeShapeCount_InRange(ByVal rng As Range) As Long
    On Error Resume Next
    SafeShapeCount_InRange = rng.ShapeRange.Count
    On Error GoTo 0
End Function

'——单元格内是否有图片（行内或浮动）
Private Function CellHasImage(ByVal c As cell) As Boolean
    CellHasImage = (c.Range.InlineShapes.Count > 0) Or (SafeShapeCount_InRange(c.Range) > 0)
End Function

'——统计“含图单元格”数量
Private Function CountImageCells(ByVal tb As Table) As Long
    Dim c As cell, N As Long
    For Each c In tb.Range.Cells
        If CellHasImage(c) Then N = N + 1
    Next c
    CountImageCells = N
End Function

'——图片表专用：表宽自适应；含图单元格→【图片格式】；其余→【图片标题】
Private Sub 格式化_图片定位表(ByVal tb As Table)
    ' 表格宽度自适应
    tb.AutoFitBehavior wdAutoFitWindow

    ' 可选：给表打上“图片定位表”的表格样式（存在就用，不存在也不报错）
    On Error Resume Next
    tb.Style = S_TABLE_PIC
    On Error GoTo 0

    ' 逐单元格设置段落样式
    Dim c As cell
    For Each c In tb.Range.Cells
        If CellHasImage(c) Then
            On Error Resume Next
            c.Range.Style = ActiveDocument.Styles(S_PARA_IMG)   ' 图片格式
            On Error GoTo 0
        Else
            On Error Resume Next
            c.Range.Style = ActiveDocument.Styles(S_PARA_CAP)   ' 图片标题
            On Error GoTo 0
        End If
    Next c
End Sub





