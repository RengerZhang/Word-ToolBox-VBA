VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocObjectInspector 
   Caption         =   "UserForm1"
   ClientHeight    =   3370
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5800
   OleObjectBlob   =   "frmDocObjectInspector.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmDocObjectInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'========================================================
'��һ��ģ���ֶΣ�������飩
'========================================================
Private mDoc As Document              ' Ŀ���ĵ�
Private mTotal As Long                ' �������
Private mIndex As Long                ' ��ǰ��ţ�1..mTotal��
Private mPos() As Long                ' ÿ������ Range.Start������
Private mInited As Boolean            ' �Ƿ��ѳ�ʼ��
Private mOverlay As Shape             ' ��ǰ����͸�����ǿ�
Private mWindowTitle As String        ' ��ǰ������⣨��ģʽ�л���

'���� ��̬ UI �ؼ��������봴������ WithEvents �����յ��¼���
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
'������һ�����Կ��أ���Ϊ False ��ȫ�־���
Private Const UI_DBG As Boolean = False
Private Const HL_DBG As Boolean = True

'��������ӡ����
Private Sub Hlog(ByVal s As String)
    If HL_DBG Then Debug.Print "[HL] " & s
End Sub

'������������ӡ��������������ߡ��߾ࡢ��ťͳһ�ߴ磩
Private Sub DumpMetrics(ByVal w As Single, ByVal h As Single, _
                        ByVal m As Single, ByVal G As Single, _
                        ByVal BW As Single, ByVal BH As Single)
    If Not UI_DBG Then Exit Sub
    Debug.Print "[UI] Metrics  W=" & w & "  H=" & h & _
                "  M=" & m & "  G=" & G & "  BW=" & BW & "  BH=" & BH
End Sub

'������������ӡ�ؼ�λ�ã���Դ������� + ������Ļ���꣩
Private Sub DumpCtrlPos(ByVal c As MSForms.Control, ByVal nm As String)
    If Not UI_DBG Then Exit Sub
    Debug.Print "[UI] " & nm & _
        "  L=" & Format(c.Left, "0.0") & _
        "  T=" & Format(c.Top, "0.0") & _
        "  W=" & Format(c.width, "0.0") & _
        "  H=" & Format(c.Height, "0.0") & _
        "  screen��(" & Format(Me.Left + c.Left, "0.0") & _
                     "," & Format(Me.Top + c.Top, "0.0") & ")"
End Sub


'======================��������������� UI����ͼ�б�������======================
Private Sub BuildUI()
    '��һ������������������λ��point��
    '    �ο�ͼ�������������ߵ� 22%����Ϣ����38%���ײ���ť����40%
    Dim w As Single, h As Single, m As Single, G As Single
    w = 190                 ' �������ͼ���ƣ�
    h = 120                ' ����ߣ���ͼ���ƣ�
    m = 8                  ' ��߾�
    G = 5                  ' �ؼ���ࣨ��/��ͳһ��

    Me.width = w + 10
    Me.Height = h
    
    '��һ�����⣺���ⲿ��ָ�������ⲿ�������Ĭ��
    If Len(mWindowTitle) = 0 Then mWindowTitle = "ȫ�ļ����"
    Me.caption = mWindowTitle


    '��������ťͳһ�ߴ磨��ͼ 3 ���ײ���ť + �Ҳࡰ��ת����ťͬ��ͬ�ߣ�
    Dim BW As Single, BH As Single
    BW = (w - 2 * m - 2 * G) / 3          ' ���ȷֵײ���
    BH = 24
    
    DumpMetrics w, h, m, G, BW, BH

    '���������⣨������壬�Ӵ֣����У�
    Set lblSummary = Me.Controls.Add("Forms.Label.1", "lblSummary", True)
    With lblSummary
        .Left = m
        .Top = m
        .width = w - 2 * m
        .Height = 15
        .caption = "���Ĺ��� 0 �� ͼƬ/���"
        .TextAlign = fmTextAlignCenter
        .Font.name = "����"
        .Font.Size = 10.5      ' ��� = 10.5pt
        .Font.bold = True
    End With

    '���ģ���Ϣ���������
    Set lblCurrent = Me.Controls.Add("Forms.Label.1", "lblCurrent", True)
    With lblCurrent
        .Left = m
        .Top = lblSummary.Top + lblSummary.Height
        .width = (w - 3 * m) - BW - G      ' �Ҳ���������ת����ť�Ŀ��
        .Height = 12
        .caption = "��ǰ�ǵ� 0 �� ͼƬ/���"
        .Font.name = "����": .Font.Size = 10.5
    End With
    
    Set lblJump = Me.Controls.Add("Forms.Label.1", "lblJump", True)
    With lblJump
        .Left = lblCurrent.Left
        .Top = lblCurrent.Top + lblCurrent.Height + G
        .width = 45
        .Height = 18
        .caption = "��ת����"
        .Font.name = "����": .Font.Size = 10.5
    End With

    Set txtJumpIndex = Me.Controls.Add("Forms.TextBox.1", "txtJumpIndex", True)
    With txtJumpIndex
        .Left = lblJump.Left + lblJump.width
        .Top = lblJump.Top - 2
        .width = 35
        .Height = 18
        .text = "1"
        .Font.name = "����": .Font.Size = 10.5
    End With

    ' �������ֱ�ǩ���ð�ʽ��ͼһ�£�
    Dim lblGe As MSForms.label
    Set lblGe = Me.Controls.Add("Forms.Label.1", "lblGe", True)
    With lblGe
        .Left = txtJumpIndex.Left + txtJumpIndex.width
        .Top = lblJump.Top - 1
        .width = 16
        .Height = 18
        .caption = "��"
        .Font.name = "����": .Font.Size = 10.5
    End With

    '���壩�Ҳࡰ��ת����ť����ײ���ť��ͬ��С��
    Set btnJump = Me.Controls.Add("Forms.CommandButton.1", "btnJump", True)
    With btnJump
        .caption = "��ת"
        .width = BW
        .Height = BH
        .Left = w - m - BW
        .Top = lblSummary.Top + lblSummary.Height + G - 2   ' ����Ϣ����ֱ����
        .TakeFocusOnClick = False
        .Font.name = "����": .Font.Size = 10.5: .Font.bold = True
        .TabIndex = 0
    End With
'    DumpCtrlPos btnJump, "btnJump"
    
    
    '�������ײ�����ť���ȿ�ȸ�
    
    Set btnPrev = Me.Controls.Add("Forms.CommandButton.1", "btnPrev", True)
    With btnPrev
        .caption = "����һ��"
        .width = BW
        .Height = BH
        .Left = m
        .Top = lblGe.Top + lblGe.Height + G
        .TakeFocusOnClick = False
        .Font.name = "����": .Font.Size = 10.5: .Font.bold = True
    End With
'    DumpCtrlPos btnPrev, "btnPrev"
    
    Set btnNext = Me.Controls.Add("Forms.CommandButton.1", "btnNext", True)
    With btnNext
        .caption = "��һ����"
        .width = BW
        .Height = BH
        .Left = btnPrev.Left + BW + G
        .Top = btnPrev.Top
        .TakeFocusOnClick = False
        .Font.name = "����": .Font.Size = 10.5: .Font.bold = True
    End With
'    DumpCtrlPos btnNext, "btnNext"
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With btnClose
        .caption = "�˳�"
        .width = BW
        .Height = BH
        .Left = btnNext.Left + BW + G
        .Top = btnPrev.Top
        .TakeFocusOnClick = False
        .Font.name = "����": .Font.Size = 10.5: .Font.bold = True
        .Cancel = True                    ' Esc �ر�
    End With
    
'    DumpCtrlPos btnClose, "btnClose"
    
End Sub

'========================================================
'�������������ڣ��Ƚ� UI���ٳ�ʼ��ҵ��
'========================================================
Private Sub UserForm_Initialize()
    BuildUI
    If Not mInited Then InitTables
End Sub

Private Sub UserForm_Terminate()
    Overlay_Clear
End Sub

'========================================================
'���ģ�������ڣ�������ʼ������ģ̬��ʾ��
'========================================================
Public Sub InitTables()
    On Error GoTo EH
    
    mWindowTitle = "ȫ�ı������"   '��һ���趨��ǰģʽ����
    Me.caption = mWindowTitle        '���������̸���һ�Σ���ʹ UI �ѽ��ã�


    '��һ����ӡ��ͼ���������ȶ�
    If ActiveWindow.View.Type <> wdPrintView Then
        ActiveWindow.View.Type = wdPrintView
    End If

    Set mDoc = ActiveDocument

    '��������������
    BuildIndex_Tables

    '�������ޱ�� �� ���ð�ť����ʾ
    If mTotal = 0 Then
        SafeEnableButtons False
        lblSummary.caption = "���Ĺ��� 0 �� ���"
        lblCurrent.caption = "��ǰ�޿ɼ�����"
        mInited = True
        Exit Sub
    End If

    '���ģ���ʼ��ţ�ѡ�����ڱ����ȣ�
    InitIndexFromSelection_Table
    If mIndex < 1 Or mIndex > mTotal Then mIndex = 1

    '���壩������ʾ
    SafeEnableButtons True
    UpdateSummaryUI
    LocateAndSelect_Table mIndex
    Overlay_TableFirstPage mDoc.Tables(mIndex)

    mInited = True
    Exit Sub
EH:
    MsgBox "��ʼ��ʧ�ܣ�" & Err.Number & " - " & Err.Description, vbExclamation
End Sub

'========================================================
'���壩����������ֻɨ�����ĵ� Tables��
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
    lblSummary.caption = "���Ĺ��� " & mTotal & " �� ���"
End Sub

'========================================================
'�������ӵ�ǰѡ���ƶ���ʼ�����
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
'���ߣ���λ������
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
'���ˣ�UI ˢ�� & ��ť״̬
'========================================================
Private Sub UpdateSummaryUI()
    lblSummary.caption = "���Ĺ��� " & mTotal & " �� ���"
    lblCurrent.caption = "��ǰ�� �� " & mIndex & " �� ���"
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
'���ţ���ť�¼�����һ/��һ/��ת/�˳���
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
        MsgBox "��������ȷ��������š�", vbExclamation: Exit Sub
    End If
    N = CLng(s)
    If N < 1 Or N > mTotal Then
        MsgBox "���볬����Χ��1 �� " & mTotal, vbExclamation: Exit Sub
    End If

    mIndex = N
    UpdateSummaryUI
    LocateAndSelect_Table mIndex
    Overlay_TableFirstPage mDoc.Tables(mIndex)
End Sub

Private Sub btnClose_Click()
    Overlay_Clear
    Me.Hide   ' ���� Unload�����ȣ�������/�������岻��������
    ��׼����ʽ������.Show
End Sub


'========================================================
'��ʮ�����Ǹ��������� D������ȫ�����������·����ޱ߿�
'========================================================
Private Sub Overlay_TableFirstPage(ByVal tbl As Table)
    On Error GoTo EH
    If ActiveWindow.View.Type <> wdPrintView Then Exit Sub  ' ֻ�ڴ�ӡ��ͼ����

    '��һ��ҳ�������ȡ�����ڽڣ�
    Dim ps As PageSetup
    Set ps = tbl.Range.Sections(1).PageSetup
    
        '����ȡ��1����Ԫ����㣨����㣩��ԡ�ҳ��/����������λ��
    Dim rCell As Range: Set rCell = tbl.cell(1, 1).Range: rCell.Collapse wdCollapseStart
    Dim xPage As Single, xText As Single, yPage As Single
    xPage = rCell.Information(wdHorizontalPositionRelativeToPage)          ' ���ҳ�����
    xText = rCell.Information(wdHorizontalPositionRelativeToTextBoundary)  ' ������������
    yPage = rCell.Information(wdVerticalPositionRelativeToPage)            ' ���ҳ���ϱ�
    
    '�����ѡ�����㡱���껻��Ϊ�������߿򡱵����꣨��ȥ�ڱ߾���߿��߿�
    Dim bwL As Single, bwT As Single
    bwL = BorderWidthPt(tbl.Borders(wdBorderLeft))   ' ��߿��߿�pt��
    bwT = BorderWidthPt(tbl.Borders(wdBorderTop))    ' �ϱ߿��߿�pt��
    
    Dim leftOuter_Page As Single, leftOuter_Text As Single, topOuter As Single
    leftOuter_Page = xPage - tbl.LeftPadding - bwL     ' ��߿���ԡ�ҳ�桱��������
    leftOuter_Text = xText - tbl.LeftPadding - bwL     ' ��߿���ԡ�����������������
    topOuter = yPage - tbl.TopPadding - bwT            ' ��߿���ԡ�ҳ�桱��������
    
    
    Dim textW As Single: textW = GetTextAreaWidth(ps)
    Dim pageH As Single: pageH = ps.PageHeight
    
    '��������������������ڣ�
    Debug.Print "[TAB] raw  xPage=" & Format(xPage, "0.0") & _
                "  xText=" & Format(xText, "0.0") & _
                "  yPage=" & Format(yPage, "0.0")
    Debug.Print "[TAB] adj  left(Page)=" & Format(leftOuter_Page, "0.0") & _
                "  left(Text)=" & Format(leftOuter_Text, "0.0") & _
                "  top(Page)=" & Format(topOuter, "0.0")
                
    '����������ֹҳ�붥����
    Dim rStart As Range, rEnd As Range
    Set rStart = tbl.Range.Duplicate: rStart.Collapse wdCollapseStart
    Set rEnd = tbl.Range.Duplicate:   rEnd.Collapse wdCollapseEnd

    Dim page0 As Long, page1 As Long
    page0 = rStart.Information(wdActiveEndAdjustedPageNumber)
    page1 = rEnd.Information(wdActiveEndAdjustedPageNumber)

    Dim topY As Single
    topY = rStart.Information(wdVerticalPositionRelativeToPage)

    '�����������꣺����ҳ�����Ʊ�β����ҳ��ҳ��ױ߾���
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

    '���ģ����Ƹ��ǿ�ê���ñ�� Range��������ձ߾ࡢ�������ҳ�棩
'    Overlay_DrawRect 0, topY, textW, (bottomY - topY), tbl.Range

    ' �ĳɣ���������߿򣺺������ҳ�棬Left �� leftOuter_Page��
    Overlay_DrawRect leftOuter_Page, topOuter, textW, (bottomY - topY), tbl.Range
    
    '������������ֱ�������pt �� cm ͬʱ��ʾ��
    Debug.Print "�����ƾ��Ρ�Left(ҳ)=" & Format(leftOuter_Page, "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(leftOuter_Page), "0.00") & "cm)��" & _
                "Top(ҳ)=" & Format(topOuter, "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(topOuter), "0.00") & "cm)��" & _
                "Width=" & Format(textW, "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(textW), "0.00") & "cm)��" & _
                "Height=" & Format((bottomY - topY), "0.0") & "pt(" & _
                Format(Application.PointsToCentimeters(bottomY - topY), "0.00") & "cm)��" & _
                "ҳ��=" & page0 & "��" & page1 & "��ê��=�� Range"
    Exit Sub
EH:
    ' ��Ĭ�����ж�������
End Sub

'���� ������Ŀ�PageWidth - ���ұ߾� - װ����
Private Function GetTextAreaWidth(ByVal ps As PageSetup) As Single
    Dim gut As Single: gut = 0
    On Error Resume Next
    If ps.Gutter > 0 Then gut = ps.Gutter
    On Error GoTo 0
    GetTextAreaWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin - gut
    If GetTextAreaWidth < 10 Then GetTextAreaWidth = 10
End Function

'���� ���ư�͸�����Σ��ޱ߿����������·���
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

    ' �ȴ�����С���Σ������ò��������꣬���� Word �ڲ��������ƫ��
    anc.Select
    Set shp = mDoc.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    With shp
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin  ' ������Ա߾�
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage        ' �������ҳ��
        .Left = 0: .Top = 0
        .width = widthPt: .Height = heightPt

        .line.Visible = msoFalse
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 236, 125)
        .Fill.Transparency = 0.45
        .WrapFormat.Type = wdWrapBehind        ' ���������·�
        .ZOrder msoSendBehindText
        .name = "InspectorOverlay"
    End With
    Set mOverlay = shp
End Sub

'���� ����ɸ��ǿ�
Private Sub Overlay_Clear()
    On Error Resume Next
    If Not mOverlay Is Nothing Then mOverlay.Delete
    Set mOverlay = Nothing
    On Error GoTo 0
End Sub

'�����ѱ߿�ö�ٿ�Ȼ��� point ֵ����������У����
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

