Attribute VB_Name = "全文_表格_完全格式化"
'====================（A）新增：按参数直接执行 ====================
Public Sub 全文表格格式化_按参数( _
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


'====================（B）核心：只做一件事——按参数跑格式化 ====================
Private Sub 全文表格格式化_核心( _
    ByVal optThick As Boolean, _
    ByVal optHeadBold As Boolean, _
    ByVal optFontPt As Single, _
    ByVal optFontName As String _
)
    '（一）变量
    Dim tb As Table
    Dim oCell As cell
    Dim myParagraph As Paragraph
    Dim n As Integer
    Dim progressForm As progressForm
    Dim i As Long, r As Long

    '（二）样式准备（字号来自参数）
    EnsureStandardTableStyle
    ActiveDocument.Styles("标准化表格样式").Font.Size = optFontPt

    '（三）计数 + 进度窗体
    i = ActiveDocument.Tables.Count
    Set progressForm = New progressForm
    progressForm.Show vbModeless
    progressForm.TextBoxStatus.text = "全文共有 " & i & " 个表格，现在开始格式化..."
    progressForm.UpdateProgressBar 0, "Processing table 1 of " & i
   

    '（四）逐表处理
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "操作已停止，正在退出..."
            Exit For
        End If

        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)

        ' 1) 应用样式 + 常规属性
        tbl.Select
        Selection.Style = "标准化表格样式"
        表格属性设置 tb

        ' 2) 内框线固定 0.5 磅
        For Each oCell In tbl.Cells
            oCell.Select
            With Selection
                .Borders.OutsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineWidth = wdLineWidth050pt
            End With

            Selection.SelectRow
            Selection.rows.AllowBreakAcrossPages = False  ' 如需跟随你原先的 enable 变量，可替换为该变量

            n = 1
            For Each myParagraph In Selection.Paragraphs
                If Len(Trim(myParagraph.Range)) = 1 Then
                    myParagraph.Range.Delete
                    n = n + 1
                End If
            Next
        Next oCell

        ' 3) 外框线：按开关 1.5 / 0.5
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = IIf(optThick, wdLineWidth150pt, wdLineWidth050pt)
            .OutsideColor = wdColorBlack
        End With

        ' 4) 首行加粗 + 每页重复：加粗由参数决定
        tbl.Select
        Selection.rows.HeadingFormat = False
        tbl.Cells(1).Select
        Selection.SelectRow
        Selection.Range.bold = optHeadBold     ' ★ 只在这里受控
        Selection.rows.HeadingFormat = True

        ' 5) 进度
        progressForm.UpdateProgressBar CLng((r / i) * 200), _
            "Processing table " & r & " of " & i
        DoEvents
    Next r

    '（五）完成提示
    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "表格格式调整完毕！"
End Sub


'====================（C）兼容保留：旧入口（弹窗取值→再调核心） ====================
Public Sub 全文表格格式化工具()
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


' （公共）表格属性设置：支持外部传入字号（磅）
Public Sub 表格属性设置(tb As Table, Optional ByVal fontPt As Single = 0!)
    '（一）自动调整 + 内边距
    tb.AutoFitBehavior (wdAutoFitWindow)
    tb.TopPadding = PixelsToPoints(0, True)
    tb.BottomPadding = PixelsToPoints(0, True)
    tb.LeftPadding = PixelsToPoints(0, True)
    tb.RightPadding = PixelsToPoints(0, True)

    '（二）准备字号：未传则从样式取；仍取不到则回退 10.5
    Dim pt As Single
    If fontPt > 0! Then
        pt = fontPt
    Else
        On Error Resume Next
        pt = ActiveDocument.Styles("标准化表格样式").Font.Size
        On Error GoTo 0
        If pt <= 0! Then pt = 10.5
    End If

    '（三）格式化表格内容
    tb.Select
    With Selection
        .rows.alignment = wdAlignRowCenter
        .rows.WrapAroundText = False

        ' 清除字体属性
        .Font.NameFarEast = ""
        .Font.NameAscii = ""
        .Range.bold = False

        ' 设置字体（字号用 pt）
        .Range.Font.NameAscii = "Times New Roman"
        .Range.Font.NameFarEast = "宋体"
        .Range.Font.Size = pt

        ' 单元格与段落设置
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        With .ParagraphFormat
            .CharacterUnitFirstLineIndent = 0
            .alignment = wdAlignParagraphCenter
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
            .LeftIndent = 0
            .RightIndent = 0
        End With

        ' 清除底色
        .Shading.BackgroundPatternColor = wdColorAutomatic

        ' 行高
        .rows.HeightRule = wdRowHeightAtLeast
        .rows.Height = CentimetersToPoints(0.6)
    End With
End Sub



