VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Լ챨�� 
   Caption         =   "XX�Լ챨��"
   ClientHeight    =   13820
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9950.001
   OleObjectBlob   =   "�Լ챨��.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "�Լ챨��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=================================================================
'=================== ���壺�Լ챨�棨ʹ�� wbReport�� ===================
Option Explicit

'=========================��һ��Win32 ���� =========================
Private Const GWL_STYLE       As Long = -16
Private Const WS_THICKFRAME   As Long = &H40000
Private Const WS_MAXIMIZEBOX  As Long = &H10000
Private Const WS_MINIMIZEBOX  As Long = &H20000

' �豸������ѯ��DPI
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

'=========================�������ṹ�� =============================
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'=========================������API ���� ===========================
#If VBA7 Then
    ' 64/32 λ��VBA7��ͳһʹ�� PtrSafe + LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetClientRect Lib "user32" ( _
        ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" ( _
        ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    ' �ϰ汾 VBA���� VBA7��
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetClientRect Lib "user32" ( _
        ByVal hWnd As Long, ByRef lpRect As RECT) As Long
    Private Declare Function GetDC Lib "user32" ( _
        ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" ( _
        ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

'=========================���ģ������¼� ===========================
' �� Activate������Ѿ�������ʱ���ÿ����� + ��������
Private Sub UserForm_Activate()
    MakeMeSizable
    ResizeReportUI
End Sub

' �� Resize�����洰�ڴ�С�仯���� wbReport
Private Sub UserForm_Resize()
    ResizeReportUI
End Sub

' �� Initialize��������� InsideWidth ��Ϊ 0������һ�γ���
Private Sub UserForm_Initialize()
    ResizeReportUI
End Sub

'=========================���壩���Ĺ��� ===========================
'��һ���Ѵ����Ϊ�ɵ���С���������/��С����ť
Private Sub MakeMeSizable()
    On Error Resume Next
#If VBA7 Then
    Dim h As LongPtr, st As LongPtr
#Else
    Dim h As Long, st As Long
#End If

    ' UserForm ʵ�ʴ�������Ϊ ThunderDFrame�����⼴ Me.Caption
    h = FindWindow("ThunderDFrame", Me.caption)
    If h <> 0 Then
        ' ��ȡ��ʽ �� ���ӿ���������/��С��ť �� д��
        #If VBA7 Then
            st = GetWindowLongPtr(h, GWL_STYLE)
            st = st Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
            Call SetWindowLongPtr(h, GWL_STYLE, st)
        #Else
            st = GetWindowLong(h, GWL_STYLE)
            st = st Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
            Call SetWindowLong(h, GWL_STYLE, st)
        #End If
    End If
End Sub

'��������ȡ����ͻ����ߴ磨�㣩�������� API��ʧ�ܻ��� InsideWidth/InsideHeight
Private Function GetClientSizePt(ByRef wPt As Single, ByRef hPt As Single) As Boolean
    On Error Resume Next
    Dim rc As RECT
#If VBA7 Then
    Dim h As LongPtr, hdc As LongPtr
#Else
    Dim h As Long, hdc As Long
#End If
    Dim dpiX As Long, dpiY As Long

    h = FindWindow("ThunderDFrame", Me.caption)
    If h = 0 Then GoTo FAIL
    If GetClientRect(h, rc) = 0 Then GoTo FAIL

    hdc = GetDC(h)
    If hdc <> 0 Then
        dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
        dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
        Call ReleaseDC(h, hdc)
    Else
        dpiX = 96: dpiY = 96
    End If

    wPt = rc.Right * 72# / dpiX
    hPt = rc.Bottom * 72# / dpiY
    GetClientSizePt = (wPt > 0 And hPt > 0)
    Exit Function
FAIL:
    GetClientSizePt = False
End Function

'������ͳһ���� wbReport ��λ����ߴ磨�����ػ��ˣ�
Private Sub ResizeReportUI()
    On Error Resume Next
    Const pad As Single = 6
    Dim cw As Single, ch As Single

    ' �� API ȡ�ͻ������㣩
    If Not GetClientSizePt(cw, ch) Then
        ' �� ���ˣ�InsideWidth/InsideHeight
        cw = Me.InsideWidth
        ch = Me.InsideHeight
        ' �� �ٻ��ˣ�Width/Height ���㣨ȥ���߿�/���⣩
        If cw <= 0 Or ch <= 0 Then
            cw = Me.width - 12          ' ���ұ߿�Լ 6*2 pt
            ch = Me.Height - 38         ' ������+���±߿��ֵ
        End If
    End If

    With Me.wbReport
        .Left = pad
        .Top = pad
        .width = cw - 2 * pad
        .Height = ch - 2 * pad
    End With
End Sub

'���ģ�ȷ�� wbReport �п�д��� Document���ȵ����� about:blank��
Private Sub EnsureBrowserDocument()
    On Error Resume Next
    If Me.wbReport.Document Is Nothing Then
        Me.wbReport.Navigate "about:blank"
        DoEvents
        Dim t As Single: t = Timer
        Do While (Me.wbReport.ReadyState <> 4) And (Timer - t < 5)
            DoEvents
        Loop
    End If
    On Error GoTo 0
End Sub


'��������������ز���Ⱦ����С�Ķ���֧�� ����/ͼ�� �����Լ죩
Public Sub LoadReportFromArray(ByRef arr As Variant, Optional ByVal captionKind As String = "��")
    ' captionKind: "��" �������飻"ͼ" ͼƬ������
    Dim isPic As Boolean: isPic = (captionKind = "ͼ")

    ' 1) �����������λ
    Me.caption = IIf(isPic, "ͼƬ�����Լ챨��", "�������Լ챨��")
    Dim unitName As String: unitName = IIf(isPic, "��ͼ", "�ű�")
    Dim idColName As String: idColName = IIf(isPic, "ͼ��", "���")

    Dim html As String, i As Long, n As Long
    n = SafeRowCount(arr)

    ' 2) ��ʽ���������ԭ��ʽ�����ַ���ƴ�ӣ�
    html = "<!doctype html><html><head><meta charset='utf-8'>" & _
           "<style>body{font-family:SimSun,'Times New Roman';font-size:10.5pt;margin:10px}" & _
           "table{border-collapse:collapse;width:100%}" & _
           "th,td{border:1px solid #e6e6e6;padding:6px 8px;vertical-align:top}" & _
           "th{background:#f4f4f4;text-align:left}" & _
           "td.col1{width:200px;white-space:nowrap}" & _
           "td.col2{width:auto}" & _
           "td.col3{width:160px;text-align:left;white-space:nowrap}" & _
           "td.col4{width:80px;text-align:center;white-space:nowrap}" & _
           ".status-ok{color:#1a7f37;font-weight:600}" & _
           ".status-red{color:#d32f2f;font-weight:600}" & _
           ".status-orange{color:#e67e22;font-weight:600}" & _
           ".status-blue{color:#1e88e5;font-weight:600}" & _
           ".bad-red{color:#d32f2f;font-weight:600}" & _
           ".bad-orange{color:#e67e22;font-weight:600}" & _
           ".bad-blue{color:#1e88e5;font-weight:600}" & _
           "a{color:#1155cc;text-decoration:underline}</style></head><body>"

    html = html & "<h3>" & IIf(isPic, "ͼƬ", "���") & "Ԥ�������" & CStr(n) & " " & unitName & "��</h3>"
    html = html & "<table><tr><th>" & idColName & "</th><th>��ǰ����͹¶�����</th><th>״̬</th><th>�༭</th></tr>"

    ' 3) �¶��Σ���һ�����գ���ɫ������λ�����֣�
    Dim orph As Variant: orph = GetOrphanRows()
    Dim ONum As Long: ONum = SafeRowCount(orph)
    If ONum > 0 Then
        Dim j As Long
        For j = 1 To ONum
            Dim raw As String:   raw = CStr(orph(j, 2))
            Dim ostart As Long:  ostart = CLng(orph(j, 3))
            Dim otext As String: otext = MidEllipsisHtml(raw, 20, 20)
            html = html & "<tr>" & _
                   "<td class='col1'>&nbsp;</td>" & _
                   "<td class='col2'><span class='bad-red'>" & otext & "</span></td>" & _
                   "<td class='col3'><span class='status-red'>" & IIf(isPic, "��ͼͷ����", "�Ǳ�ͷ����") & "</span></td>" & _
                   "<td class='col4'><a href='cmd?edit=" & CStr(ostart) & "'>��λ</a></td>" & _
                   "</tr>"
        Next
    End If

    ' 4) �����У���ԭ�߼���ͬ��ֻ�ѡ����ж�/���ֻ��� captionKind��
    For i = 1 To n
        Dim label As String:  label = CStr(arr(i, 12))    ' ����X-X����ͼX-X��
        Dim rawT As String:   rawT = CStr(arr(i, 4))      ' �������ڶ�ԭ��
        Dim isCap As Boolean: isCap = (arr(i, 6) = True)  ' �Ƿ�����Ŀ����ʽ
        Dim tStart As Long:   tStart = CLng(arr(i, 2))    ' ��Ӧ������㣨��/ͼ��
        Dim pstart As Long:   pstart = CLng(arr(i, 3))    ' ��������

        Dim col2Txt As String
        If Len(rawT) > 150 Then
            col2Txt = HtmlEncode(Left$(rawT, 150)) & "��"
        Else
            col2Txt = HtmlEncode(rawT)
        End If

        Dim status As String
        If Not isCap Then
            status = "<span class='status-orange'>" & captionKind & "������ʽ����</span>"
            col2Txt = "<span class='bad-orange'>" & col2Txt & "</span>"
        Else
            Dim firstCh As String: firstCh = FirstVisibleChar(rawT)
            If firstCh <> captionKind Then
                status = "<span class='status-blue'>" & captionKind & "�����Ŵ���</span>"
                col2Txt = "<span class='bad-blue'>" & col2Txt & "</span>"
            Else
                status = "<span class='status-ok'>" & captionKind & "�����ʽ��ȷ</span>"
            End If
        End If

        html = html & "<tr>" & _
               "<td class='col1'><a href='cmd?goto=" & CStr(tStart) & "'>" & HtmlEncode(label) & "</a></td>" & _
               "<td class='col2'>" & col2Txt & "</td>" & _
               "<td class='col3'>" & status & "</td>" & _
               "<td class='col4'>" & IIf(pstart > 0, "<a href='cmd?edit=" & CStr(pstart) & "'>�༭</a>", "<span class='status-red'>��</span>") & "</td>" & _
               "</tr>"
    Next

    html = html & "</table></body></html>"

    EnsureBrowserDocument
    Me.wbReport.Document.Open
    Me.wbReport.Document.Write html
    Me.wbReport.Document.Close
End Sub




'���ģ����س����Ӳ���λ���ؼ��¼�����Ϊ wbReport_BeforeNavigate2��
Private Sub wbReport_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, _
        Flags As Variant, TargetFrameName As Variant, PostData As Variant, _
        Headers As Variant, Cancel As Boolean)
    
    On Error GoTo SAFE_EXIT
    Dim s As String: s = CStr(URL)
    
    '��һ��ͬʱ���ݣ�about:cmd?goto=xxxx �Լ� cmd?goto=xxxx
    Dim hitPos As Long
    hitPos = InStr(1, s, "cmd?goto=", vbTextCompare)
    If hitPos = 0 Then hitPos = InStr(1, s, "about:cmd?goto=", vbTextCompare)
    
    If hitPos > 0 Then
        Cancel = True
        Dim v As Long
        v = CLng(val(mid$(s, hitPos + Len("cmd?goto=")))) ' Val ���������Զ�ֹͣ
        GoToDocumentPos v
    End If
    
    ' �����������༭��ڣ��������ö���㣬����ֱ���޸�
    Dim hitE As Long: hitE = InStr(1, s, "cmd?edit=", vbTextCompare)
    If hitE > 0 Then
        Cancel = True
        Dim epos As Long: epos = CLng(val(mid$(s, hitE + Len("cmd?edit="))))
        GoToDocumentPos epos
        Exit Sub
    End If

SAFE_EXIT:
End Sub

'���壩��λ��ָ�� Range.Start
Private Sub GoToDocumentPos(ByVal startPos As Long)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    Dim r As Range: Set r = doc.Range(Start:=startPos, End:=startPos)
    r.Select
    If Err.Number = 0 Then
        Application.Activate
        ActiveWindow.ScrollIntoView r, True
    End If
    On Error GoTo 0
End Sub

'������HTML ת��
Private Function HtmlEncode(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    s = Replace$(s, """", "&quot;")
    HtmlEncode = s
End Function

'���ߣ����ɡ���ǰleftN + �� + �κ�rightN����HTML��ȫժҪ
Private Function MidEllipsisHtml(ByVal s As String, ByVal leftN As Long, ByVal rightN As Long) As String
    Dim L As Long: L = Len(s)
    If L <= leftN + rightN Then
        MidEllipsisHtml = HtmlEncode(s)
    Else
        Dim leftPart As String:  leftPart = Left$(s, leftN)
        Dim rightPart As String: rightPart = Right$(s, rightN)
        MidEllipsisHtml = HtmlEncode(leftPart) & "&hellip;" & HtmlEncode(rightPart)
    End If
End Function
'����ȡ���׸��ɼ��ַ�����ȥ�����С���Ԫ���������ȫ�ǿո��� Trim��
Private Function FirstVisibleChar(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = Trim$(s)
    If Len(s) > 0 Then
        FirstVisibleChar = Left$(s, 1)
    Else
        FirstVisibleChar = ""
    End If
End Function

'��һ����ȫ��ȡ��ά�������������� UBound �����鱨��
Private Function SafeRowCount(ByVal v As Variant) As Long
    On Error GoTo FAIL
    If IsArray(v) Then
        SafeRowCount = UBound(v, 1) - LBound(v, 1) + 1
        Exit Function
    End If
FAIL:
    SafeRowCount = 0
End Function
