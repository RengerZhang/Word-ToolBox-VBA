VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 自检报告 
   Caption         =   "XX自检报告"
   ClientHeight    =   13820
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9950.001
   OleObjectBlob   =   "自检报告.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "自检报告"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=================================================================
'=================== 窗体：自检报告（使用 wbReport） ===================
Option Explicit

'=========================（一）Win32 常量 =========================
Private Const GWL_STYLE       As Long = -16
Private Const WS_THICKFRAME   As Long = &H40000
Private Const WS_MAXIMIZEBOX  As Long = &H10000
Private Const WS_MINIMIZEBOX  As Long = &H20000

' 设备能力查询：DPI
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

'=========================（二）结构体 =============================
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'=========================（三）API 声明 ===========================
#If VBA7 Then
    ' 64/32 位（VBA7）统一使用 PtrSafe + LongPtr
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
    ' 老版本 VBA（非 VBA7）
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

'=========================（四）公共事件 ===========================
' ① Activate：句柄已就绪，此时启用可拉伸 + 调整布局
Private Sub UserForm_Activate()
    MakeMeSizable
    ResizeReportUI
End Sub

' ② Resize：跟随窗口大小变化铺满 wbReport
Private Sub UserForm_Resize()
    ResizeReportUI
End Sub

' ③ Initialize：个别机器 InsideWidth 仍为 0；先做一次尝试
Private Sub UserForm_Initialize()
    ResizeReportUI
End Sub

'=========================（五）核心功能 ===========================
'（一）把窗体改为可调大小并启用最大/最小化按钮
Private Sub MakeMeSizable()
    On Error Resume Next
#If VBA7 Then
    Dim h As LongPtr, st As LongPtr
#Else
    Dim h As Long, st As Long
#End If

    ' UserForm 实际窗口类名为 ThunderDFrame，标题即 Me.Caption
    h = FindWindow("ThunderDFrame", Me.caption)
    If h <> 0 Then
        ' 读取样式 → 叠加可拉伸和最大/最小按钮 → 写回
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

'（二）获取窗体客户区尺寸（点），优先用 API；失败回退 InsideWidth/InsideHeight
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

'（三）统一调整 wbReport 的位置与尺寸（含多重回退）
Private Sub ResizeReportUI()
    On Error Resume Next
    Const pad As Single = 6
    Dim cw As Single, ch As Single

    ' ① API 取客户区（点）
    If Not GetClientSizePt(cw, ch) Then
        ' ② 回退：InsideWidth/InsideHeight
        cw = Me.InsideWidth
        ch = Me.InsideHeight
        ' ③ 再回退：Width/Height 估算（去掉边框/标题）
        If cw <= 0 Or ch <= 0 Then
            cw = Me.width - 12          ' 左右边框约 6*2 pt
            ch = Me.Height - 38         ' 标题栏+上下边框估值
        End If
    End If

    With Me.wbReport
        .Left = pad
        .Top = pad
        .width = cw - 2 * pad
        .Height = ch - 2 * pad
    End With
End Sub

'（四）确保 wbReport 有可写入的 Document（先导航到 about:blank）
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


'（三）从数组加载并渲染（最小改动：支持 “表/图” 两种自检）
Public Sub LoadReportFromArray(ByRef arr As Variant, Optional ByVal captionKind As String = "表")
    ' captionKind: "表" 表格标题检查；"图" 图片标题检查
    Dim isPic As Boolean: isPic = (captionKind = "图")

    ' 1) 标题与计数单位
    Me.caption = IIf(isPic, "图片标题自检报告", "表格标题自检报告")
    Dim unitName As String: unitName = IIf(isPic, "张图", "张表")
    Dim idColName As String: idColName = IIf(isPic, "图号", "表号")

    Dim html As String, i As Long, n As Long
    n = SafeRowCount(arr)

    ' 2) 样式（保持你的原样式，仅字符串拼接）
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

    html = html & "<h3>" & IIf(isPic, "图片", "表格") & "预检查结果（" & CStr(n) & " " & unitName & "）</h3>"
    html = html & "<table><tr><th>" & idColName & "</th><th>表前段落和孤儿段落</th><th>状态</th><th>编辑</th></tr>"

    ' 3) 孤儿段（第一列留空；红色；“定位”保持）
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
                   "<td class='col3'><span class='status-red'>" & IIf(isPic, "非图头段落", "非表头段落") & "</span></td>" & _
                   "<td class='col4'><a href='cmd?edit=" & CStr(ostart) & "'>定位</a></td>" & _
                   "</tr>"
        Next
    End If

    ' 4) 主体行（与原逻辑相同，只把“表”判断/文字换成 captionKind）
    For i = 1 To n
        Dim label As String:  label = CStr(arr(i, 12))    ' “表X-X”或“图X-X”
        Dim rawT As String:   rawT = CStr(arr(i, 4))      ' 标题所在段原文
        Dim isCap As Boolean: isCap = (arr(i, 6) = True)  ' 是否套了目标样式
        Dim tStart As Long:   tStart = CLng(arr(i, 2))    ' 对应对象起点（表/图）
        Dim pstart As Long:   pstart = CLng(arr(i, 3))    ' 标题段起点

        Dim col2Txt As String
        If Len(rawT) > 150 Then
            col2Txt = HtmlEncode(Left$(rawT, 150)) & "…"
        Else
            col2Txt = HtmlEncode(rawT)
        End If

        Dim status As String
        If Not isCap Then
            status = "<span class='status-orange'>" & captionKind & "标题样式错误</span>"
            col2Txt = "<span class='bad-orange'>" & col2Txt & "</span>"
        Else
            Dim firstCh As String: firstCh = FirstVisibleChar(rawT)
            If firstCh <> captionKind Then
                status = "<span class='status-blue'>" & captionKind & "标题编号错误</span>"
                col2Txt = "<span class='bad-blue'>" & col2Txt & "</span>"
            Else
                status = "<span class='status-ok'>" & captionKind & "标题格式正确</span>"
            End If
        End If

        html = html & "<tr>" & _
               "<td class='col1'><a href='cmd?goto=" & CStr(tStart) & "'>" & HtmlEncode(label) & "</a></td>" & _
               "<td class='col2'>" & col2Txt & "</td>" & _
               "<td class='col3'>" & status & "</td>" & _
               "<td class='col4'>" & IIf(pstart > 0, "<a href='cmd?edit=" & CStr(pstart) & "'>编辑</a>", "<span class='status-red'>—</span>") & "</td>" & _
               "</tr>"
    Next

    html = html & "</table></body></html>"

    EnsureBrowserDocument
    Me.wbReport.Document.Open
    Me.wbReport.Document.Write html
    Me.wbReport.Document.Close
End Sub




'（四）拦截超链接并定位（控件事件名改为 wbReport_BeforeNavigate2）
Private Sub wbReport_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, _
        Flags As Variant, TargetFrameName As Variant, PostData As Variant, _
        Headers As Variant, Cancel As Boolean)
    
    On Error GoTo SAFE_EXIT
    Dim s As String: s = CStr(URL)
    
    '（一）同时兼容：about:cmd?goto=xxxx 以及 cmd?goto=xxxx
    Dim hitPos As Long
    hitPos = InStr(1, s, "cmd?goto=", vbTextCompare)
    If hitPos = 0 Then hitPos = InStr(1, s, "about:cmd?goto=", vbTextCompare)
    
    If hitPos > 0 Then
        Cancel = True
        Dim v As Long
        v = CLng(val(mid$(s, hitPos + Len("cmd?goto=")))) ' Val 遇非数字自动停止
        GoToDocumentPos v
    End If
    
    ' ——新增：编辑入口，先跳到该段起点，方便直接修改
    Dim hitE As Long: hitE = InStr(1, s, "cmd?edit=", vbTextCompare)
    If hitE > 0 Then
        Cancel = True
        Dim epos As Long: epos = CLng(val(mid$(s, hitE + Len("cmd?edit="))))
        GoToDocumentPos epos
        Exit Sub
    End If

SAFE_EXIT:
End Sub

'（五）定位到指定 Range.Start
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

'（六）HTML 转义
Private Function HtmlEncode(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    s = Replace$(s, """", "&quot;")
    HtmlEncode = s
End Function

'（七）生成“段前leftN + … + 段后rightN”的HTML安全摘要
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
'——取“首个可见字符”（去掉换行、单元格结束符、全角空格，再 Trim）
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

'（一）安全获取二维数组行数，避免 UBound 空数组报错
Private Function SafeRowCount(ByVal v As Variant) As Long
    On Error GoTo FAIL
    If IsArray(v) Then
        SafeRowCount = UBound(v, 1) - LBound(v, 1) + 1
        Exit Function
    End If
FAIL:
    SafeRowCount = 0
End Function
