VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocObjectInspector 
   Caption         =   "UserForm1"
   ClientHeight    =   3370
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5800
   OleObjectBlob   =   "frmDocObjectInspector.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmDocObjectInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'========================================================
'（一）模块字段（仅表格检查）
'========================================================
Private mDoc As Document              ' 目标文档
Private mTotal As Long                ' 表格总数
Private mIndex As Long                ' 当前序号（1..mTotal）
Private mPos() As Long                ' 每个表格的 Range.Start（升序）
Private mInited As Boolean            ' 是否已初始化
Private mOverlay As Shape             ' 当前“半透明覆盖框”
Private mWindowTitle As String        ' 当前窗体标题（随模式切换）

'—— 动态 UI 控件（纯代码创建，需 WithEvents 才能收到事件）
Private WithEvents btnPrev  As MSForms.CommandButton
Attribute btnPrev.VB_VarHelpID = -1
Private WithEvents btnNext  As MSForms.CommandButton
Attribute btnNext.VB_VarHelpID = -1
Private WithEvents btnJump  As MSForms.CommandButton
Attribute btnJump.VB_VarHelpID = -1
Private WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1
Private lblSummary   As MSForms.label
Private lblCurrent   As MSForms.label
Private lblJump      As MSForms.label
Private txtJumpIndex As MSForms.TextBox
'——（一）调试开关：改为 False 可全局静音
Private Const UI_DBG As Boolean = False
Private Const HL_DBG As Boolean = True

'（二）打印助手
Private Sub Hlog(ByVal s As String)
    If HL_DBG Then Debug.Print "[HL] " & s
End Sub

'——（二）打印整体度量（窗体宽高、边距、按钮统一尺寸）
Private Sub DumpMetrics(ByVal w As Single, ByVal h As Single, _
                        ByVal m As Single, ByVal G As Single, _
                        ByVal BW As Single, ByVal BH As Single)
    If Not UI_DBG Then Exit Sub
    Debug.Print "[UI] Metrics  W=" & w & "  H=" & h & _
                "  M=" & m & "  G=" & G & "  BW=" & BW & "  BH=" & BH
End Sub

'——（三）打印控件位置（相对窗体坐标 + 估算屏幕坐标）
Private Sub DumpCtrlPos(ByVal c As MSForms.Control, ByVal nm As String)
    If Not UI_DBG Then Exit Sub
    Debug.Print "[UI] " & nm & _
        "  L=" & Format(c.Left, "0.0") & _
        "  T=" & Format(c.Top, "0.0") & _
        "  W=" & Format(c.width, "0.0") & _
        "  H=" & Format(c.Height, "0.0") & _
        "  screen≈(" & Format(Me.Left + c.Left, "0.0") & _
                     "," & Format(Me.Top + c.Top, "0.0") & ")"
End Sub


'======================（二）纯代码绘制 UI（按图中比例与规格）======================
Private Sub BuildUI()
    '（一）整体比例与度量（单位：point）
    '    参考图：标题区≈整高的 22%，信息区≈38%，底部按钮区≈40%
    Dim w As Single, h As Single, m As Single, G As Single
    w = 190                 ' 窗体宽（按图近似）
    h = 120                ' 窗体高（按图近似）
    m = 8                  ' 外边距
    G = 5                  ' 控件间距（横/纵统一）

    Me.width = w + 10
    Me.Height = h
    
    '（一）标题：若外部已指定则用外部，否则给默认
    If Len(mWindowTitle) = 0 Then mWindowTitle = "全文检查器"
    Me.caption = mWindowTitle


    '（二）按钮统一尺寸（按图 3 个底部按钮 + 右侧“跳转”按钮同宽同高）
    Dim BW As Single, BH As Single
    BW = (w - 2 * m - 2 * G) / 3          ' 三等分底部区
    BH = 24
    
    DumpMetrics w, h, m, G, BW, BH

    '（三）标题（五号宋体，加粗，居中）
    Set lblSummary = Me.Controls.Add("Forms.Label.1", "lblSummary", True)
    With lblSummary
        .Left = m
        .Top = m
        .width = w - 2 * m
        .Height = 15
        .caption = "本文共有 0 个 图片/表格"
        .TextAlign = fmTextAlignCenter
        .Font.name = "宋体"
        .Font.Size = 10.5      ' 五号 = 10.5pt
        .Font.bold = True
    End With

    '（四）信息区左侧两行
    Set lblCurrent = Me.Controls.Add("Forms.Label.1", "lblCurrent", True)
    With lblCurrent
        .Left = m
        .Top = lblSummary.Top + lblSummary.Height
        .width = (w - 3 * m) - BW - G      ' 右侧留给“跳转”按钮的宽度
        .Height = 12
        .caption = "当前是第 0 个 图片/表格"
        .Font.name = "宋体": .Font.Size = 10.5
    End With
    
    Set lblJump = Me.Controls.Add("Forms.Label.1", "lblJump", True)
    With lblJump
        .Left = lblCurrent.Left
        .Top = lblCurrent.Top + lblCurrent.Height + G
        .width = 45
        .Height = 18
        .caption = "跳转到第"
        .Font.name = "宋体": .Font.Size = 10.5
    End With

    Set txtJumpIndex = Me.Controls.Add("Forms.TextBox.1", "txtJumpIndex", True)
    With txtJumpIndex
        .Left = lblJump.Left + lblJump.width
        .Top = lblJump.Top - 2
        .width = 35
        .Height = 18
        .text = "1"
        .Font.name = "宋体": .Font.Size = 10.5
    End With

    ' “个”字标签（让版式与图一致）
    Dim lblGe As MSForms.label
    Set lblGe = Me.Controls.Add("Forms.Label.1", "lblGe", True)
    With lblGe
        .Left = txtJumpIndex.Left + txtJumpIndex.width
        .Top = lblJump.Top - 1
        .width = 16
        .Height = 18
        .caption = "个"
        .Font.name = "宋体": .Font.Size = 10.5
    End With

    '（五）右侧“跳转”按钮（与底部按钮相同大小）
    Set btnJump = Me.Controls.Add("Forms.CommandButton.1", "btnJump", True)
    With btnJump
        .caption = "跳转"
        .width = BW
        .Height = BH
        .Left = w - m - BW
        .Top = lblSummary.Top + lblSummary.Height + G - 2   ' 与信息区垂直对齐
        .TakeFocusOnClick = False
        .Font.name = "宋体": .Font.Size = 10.5: .Font.bold = True
        .TabIndex = 0
    End With
'    DumpCtrlPos btnJump, "btnJump"
    
    
    '（六）底部三按钮：等宽等高
    
    Set btnPrev = Me.Controls.Add("Forms.CommandButton.1", "btnPrev", True)
    With btnPrev
        .caption = "←上一个"
        .width = BW
        .Height = BH
        .Left = m
        .Top = lblGe.Top + lblGe.Height + G
        .TakeFocusOnClick = False
        .Font.name = "宋体": .Font.Size = 10.5: .Font.bold = True
    End With
'    DumpCtrlPos btnPrev, "btnPrev"
    
    Set btnNext = Me.Controls.Add("Forms.CommandButton.1", "btnNext", True)
    With btnNext
        .caption = "下一个→"
        .width = BW
        .Height = BH
        .Left = btnPrev.Left + BW + G
        .Top = btnPrev.Top
        .TakeFocusOnClick = False
        .Font.name = "宋体": .Font.Size = 10.5: .Font.bold = True
    End With
'    DumpCtrlPos btnNext, "btnNext"
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With btnClose
        .caption = "退出"
        .width = BW
        .Height = BH
        .Left = btnNext.Left + BW + G
        .Top = btnPrev.Top
        .TakeFocusOnClick = False
        .Font.name = "宋体": .Font.Size = 10.5: .Font.bold = True
        .Cancel = True                    ' Esc 关闭
    End With
    
'    DumpCtrlPos btnClose, "btnClose"
    
End Sub

'========================================================
'（三）生命周期：先建 UI，再初始化业务
'========================================================
Private Sub UserForm_Initialize()
    BuildUI
    If Not mInited Then InitTables
End Sub

Private Sub UserForm_Terminate()
    Overlay_Clear
End Sub

'========================================================
'（四）对外入口：仅表格初始化（非模态显示）
'========================================================
Public Sub InitTables()
    On Error GoTo EH
    
    mWindowTitle = "全文表格检查器"   '（一）设定当前模式标题
    Me.caption = mWindowTitle        '（二）立刻覆盖一次（即使 UI 已建好）


    '（一）打印视图下坐标最稳定
    If ActiveWindow.View.Type <> wdPrintView Then
        ActiveWindow.View.Type = wdPrintView
    End If

    Set mDoc = ActiveDocument

    '（二）构建索引
    BuildIndex_Tables

    '（三）无表格 → 禁用按钮并提示
    If mTotal = 0 Then
        SafeEnableButtons False
        lblSummary.caption = "本文共有 0 个 表格"
        lblCurrent.caption = "当前无可检查对象"
        mInited = True
        Exit Sub
    End If

    '（四）起始序号（选区所在表优先）
    InitIndexFromSelection_Table
    If mIndex < 1 Or mIndex > mTotal Then mIndex = 1

    '（五）首屏显示
    SafeEnableButtons True
    UpdateSummaryUI
    LocateAndSelect_Table mIndex
    Overlay_TableFirstPage mDoc.Tables(mIndex)

    mInited = True
    Exit Sub
EH:
    MsgBox "初始化失败：" & Err.Number & " - " & Err.Description, vbExclamation
End Sub

'========================================================
'（五）索引构建（只扫描正文的 Tables）
'========================================================
Private Sub BuildIndex_Tables()
    Dim i As Long
    mTotal = mDoc.Tables.Count
    If mTotal > 0 Then
        ReDim mPos(1 To mTotal)
        For i = 1 To mTotal
            mPos(i) = mDoc.Tables(i).Range.Start
        Next
    Else
        Erase mPos
    End If
    lblSummary.caption = "本文共有 " & mTotal & " 个 表格"
End Sub

'========================================================
'（六）从当前选区推断起始表序号
'========================================================
Private Sub InitIndexFromSelection_Table()
    Dim s As Long: s = Selection.Range.Start
    Dim i As Long

    If Selection.Information(wdWithInTable) Then
        For i = 1 To mDoc.Tables.Count
            If s >= mDoc.Tables(i).Range.Start And s < mDoc.Tables(i).Range.End Then
                mIndex = i: Exit Sub
            End If
        Next
    End If

    For i = 1 To mTotal
        If mPos(i) >= s Then mIndex = i: Exit Sub
    Next
    mIndex = mTotal
End Sub

'========================================================
'（七）定位并滚动
'========================================================
Private Sub LocateAndSelect_Table(ByVal idx As Long)
    On Error GoTo CLEAN
    Application.ScreenUpdating = False

    If mTotal = 0 Then GoTo CLEAN
    If idx < 1 Then idx = 1
    If idx > mTotal Then idx = mTotal

    mDoc.Tables(idx).Range.Select
    ActiveWindow.ScrollIntoView Selection.Range, True

CLEAN:
    Application.ScreenUpdating = True
End Sub

'========================================================
'（八）UI 刷新 & 按钮状态
'========================================================
Private Sub UpdateSummaryUI()
    lblSummary.caption = "本文共有 " & mTotal & " 个 表格"
    lblCurrent.caption = "当前是 第 " & mIndex & " 个 表格"
    Me.Repaint
End Sub

Private Sub SafeEnableButtons(ByVal yes As Boolean)
    On Error Resume Next
    btnPrev.Enabled = yes
    btnNext.Enabled = yes
    btnJump.Enabled = yes
    On Error GoTo 0
End Sub

'========================================================
'（九）按钮事件（上一/下一/跳转/退出）
'========================================================
Private Sub btnPrev_Click()
    If Not mInited Then InitTables
    If mTotal = 0 Then Exit Sub
    mIndex = mIndex - 1: If mIndex < 1 Then mIndex = mTotal
    UpdateSummaryUI
    LocateAndSelect_Table mIndex
    Overlay_TableFirstPage mDoc.Tables(mIndex)
End Sub

Private Sub btnNext_Click()
    If Not mInited Then InitTables
    If mTotal = 0 Then Exit Sub
    mIndex = mIndex + 1: If mIndex > mTotal Then mIndex = 1
    UpdateSummaryUI
    LocateAndSelect_Table mIndex
    Overlay_TableFirstPage mDoc.Tables(mIndex)
End Sub

Private Sub btnJump_Click()
    If Not mInited Then InitTables
    If mTotal = 0 Then Exit Sub

    Dim s As String, N As Long
    s = Trim$(txtJumpIndex.text)
    If Len(s) = 0 Or Not IsNumeric(s) Then
        MsgBox "请输入正确的数字序号。", vbExclamation: Exit Sub
    End If
    N = CLng(s)
    If N < 1 Or N > mTotal Then
        MsgBox "输入超出范围：1 ～ " & mTotal, vbExclamation: Exit Sub
    End If

    mIndex = N
    UpdateSummaryUI
    LocateAndSelect_Table mIndex
    Overlay_TableFirstPage mDoc.Tables(mIndex)
End Sub

Private Sub btnClose_Click()
    Overlay_Clear
    Me.Hide   ' 不用 Unload，更稳（主窗口/其它窗体不被阻塞）
    标准化格式工具箱.Show
End Sub


'========================================================
'（十）覆盖高亮（方案 D：版心全宽、衬于文字下方、无边框）
'========================================================
Private Sub Overlay_TableFirstPage(ByVal tbl As Table)
    On Error GoTo EH
    If ActiveWindow.View.Type <> wdPrintView Then Exit Sub  ' 只在打印视图绘制

    '（一）页面参数（取表所在节）
    Dim ps As PageSetup
    Set ps = tbl.Range.Sections(1).PageSetup
    
        '——取第1个单元格起点（插入点）相对“页面/文字区”的位置
    Dim rCell As Range: Set rCell = tbl.cell(1, 1).Range: rCell.Collapse wdCollapseStart
    Dim xPage As Single, xText As Single, yPage As Single
    xPage = rCell.Information(wdHorizontalPositionRelativeToPage)          ' 相对页面左边
    xText = rCell.Information(wdHorizontalPositionRelativeToTextBoundary)  ' 相对文字区左边
    yPage = rCell.Information(wdVerticalPositionRelativeToPage)            ' 相对页面上边
    
    '——把“插入点”坐标换算为“表格外边框”的坐标（减去内边距与边框线宽）
    Dim bwL As Single, bwT As Single
    bwL = BorderWidthPt(tbl.Borders(wdBorderLeft))   ' 左边框线宽（pt）
    bwT = BorderWidthPt(tbl.Borders(wdBorderTop))    ' 上边框线宽（pt）
    
    Dim leftOuter_Page As Single, leftOuter_Text As Single, topOuter As Single
    leftOuter_Page = xPage - tbl.LeftPadding - bwL     ' 外边框相对“页面”的左坐标
    leftOuter_Text = xText - tbl.LeftPadding - bwL     ' 外边框相对“文字区”的左坐标
    topOuter = yPage - tbl.TopPadding - bwT            ' 外边框相对“页面”的上坐标
    
    
    Dim textW As Single: textW = GetTextAreaWidth(ps)
    Dim pageH As Single: pageH = ps.PageHeight
    
    '——调试输出（立即窗口）
    Debug.Print "[TAB] raw  xPage=" & Format(xPage, "0.0") & _
                "  xText=" & Format(xText, "0.0") & _
                "  yPage=" & Format(yPage, "0.0")
    Debug.Print "[TAB] adj  left(Page)=" & Format(leftOuter_Page, "0.0") & _
                "  left(Text)=" & Format(leftOuter_Text, "0.0") & _
                "  top(Page)=" & Format(topOuter, "0.0")
                
    '（二）表起止页与顶坐标
    Dim rStart As Range, rEnd As Range
    Set rStart = tbl.Range.Duplicate: rStart.Collapse wdCollapseStart
    Set rEnd = tbl.Range.Duplicate:   rEnd.Collapse wdCollapseEnd

    Dim page0 As Long, page1 As Long
    page0 = rStart.Information(wdActiveEndAdjustedPageNumber)
    page1 = rEnd.Information(wdActiveEndAdjustedPageNumber)

    Dim topY As Single
    topY = rStart.Information(wdVerticalPositionRelativeToPage)

    '（三）底坐标：不跨页→近似表尾；跨页→页面底边距线
    Dim bottomY As Single
    If page1 = page0 Then
        Dim tailY As Single: tailY = rEnd.Information(wdVerticalPositionRelativeToPage)
        Dim lastRow As row, lastTop As Single, lastEst As Single
        Set lastRow = tbl.rows(tbl.rows.Count)
        lastTop = lastRow.Cells(1).Range.Information(wdVerticalPositionRelativeToPage)

        If lastRow.HeightRule <> wdRowHeightAuto And lastRow.Height > 0 Then
            lastEst = lastTop + lastRow.Height
        Else
            lastEst = tailY + 1
        End If
        bottomY = IIf(lastEst > tailY, lastEst, tailY)
        If bottomY <= topY Then bottomY = topY + 2
    Else
        bottomY = pageH - ps.BottomMargin
    End If

    '（四）绘制覆盖框（锚点用表格 Range；横向参照边距、纵向参照页面）
'    Overlay_DrawRect 0, topY, textW, (bottomY - topY), tbl.Range

    ' 改成（对齐表格外边框：横向参照页面，Left 用 leftOuter_Page）
    Overlay_DrawRect leftOuter_Page, topOuter, textW, (bottomY - topY), tbl.Range
    
    '——立即窗口直观输出（pt 与 cm 同时显示）
    Debug.Print "【绘制矩形】Left(页)=" & Format(leftOuter_Page, "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(leftOuter_Page), "0.00") & "cm)，" & _
                "Top(页)=" & Format(topOuter, "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(topOuter), "0.00") & "cm)，" & _
                "Width=" & Format(textW, "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(textW), "0.00") & "cm)，" & _
                "Height=" & Format((bottomY - topY), "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(bottomY - topY), "0.00") & "cm)；" & _
                "页码=" & page0 & "→" & page1 & "；锚点=表 Range"
    Exit Sub
EH:
    ' 静默，不中断主流程
End Sub

'—— 计算版心宽：PageWidth - 左右边距 - 装订线
Private Function GetTextAreaWidth(ByVal ps As PageSetup) As Single
    Dim gut As Single: gut = 0
    On Error Resume Next
    If ps.Gutter > 0 Then gut = ps.Gutter
    On Error GoTo 0
    GetTextAreaWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin - gut
    If GetTextAreaWidth < 10 Then GetTextAreaWidth = 10
End Function

'—— 绘制半透明矩形（无边框，置于文字下方）
Private Sub Overlay_DrawRect(ByVal leftPt As Single, ByVal topPt As Single, _
                             ByVal widthPt As Single, ByVal heightPt As Single, _
                             Optional ByVal anchorRng As Range = Nothing)

    Dim shp As Shape, anc As Range
    Overlay_Clear

    If anchorRng Is Nothing Then
        Set anc = Selection.Range
    Else
        Set anc = anchorRng
    End If

    ' 先创建最小矩形，再设置参照与坐标，避免 Word 内部换算造成偏移
    anc.Select
    Set shp = mDoc.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    With shp
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin  ' 横向相对边距
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage        ' 纵向相对页面
        .Left = 0: .Top = 0
        .width = widthPt: .Height = heightPt

        .line.Visible = msoFalse
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 236, 125)
        .Fill.Transparency = 0.45
        .WrapFormat.Type = wdWrapBehind        ' 衬于文字下方
        .ZOrder msoSendBehindText
        .name = "InspectorOverlay"
    End With
    Set mOverlay = shp
End Sub

'—— 清除旧覆盖框
Private Sub Overlay_Clear()
    On Error Resume Next
    If Not mOverlay Is Nothing Then mOverlay.Delete
    Set mOverlay = Nothing
    On Error GoTo 0
End Sub

'——把边框枚举宽度换成 point 值（用于坐标校正）
Private Function BorderWidthPt(ByVal b As Border) As Single
    Select Case b.LineWidth
        Case wdLineWidth025pt: BorderWidthPt = 0.25
        Case wdLineWidth050pt: BorderWidthPt = 0.5
        Case wdLineWidth075pt: BorderWidthPt = 0.75
        Case wdLineWidth100pt: BorderWidthPt = 1#
        Case wdLineWidth150pt: BorderWidthPt = 1.5
        Case wdLineWidth225pt: BorderWidthPt = 2.25
        Case wdLineWidth300pt: BorderWidthPt = 3#
        Case wdLineWidth450pt: BorderWidthPt = 4.5
        Case wdLineWidth600pt: BorderWidthPt = 6#
        Case Else:              BorderWidthPt = 0#
    End Select
End Function

