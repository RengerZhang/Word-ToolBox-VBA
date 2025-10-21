VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 标准化格式工具箱 
   Caption         =   "标准化格式工具箱  V1.0.250919"
   ClientHeight    =   5270
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9170.001
   OleObjectBlob   =   "标准化格式工具箱.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "标准化格式工具箱"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' —— 窗体返回值（主程序 Show 后读取）——
Public SelectedThickOuter As Boolean      ' 外框1.5磅开关（全文）
Public SelectedFirstRowBold As Boolean    ' 首行加粗开关（全文）
Public SelectedFontSizeName As String     ' 中文字号名（全文）
Public SelectedFontSizePt As Single       ' 对应磅值（全文）
Public Canceled As Boolean                ' 是否取消


' =========================
'  窗体初始化（全量初始化）
' =========================
Private Sub UserForm_Initialize()
    '（一）全部子页面一次性初始化（强制重置默认）
    Init_All Me, True

    '（二）【要求】表格格式化页的控件默认值已移动到“初始化中心”，
    '     你原来在这里对 chk/cbo 的默认赋值可以删除，避免重复。
End Sub

' =========================
'  通用小工具
' =========================
'（一）判断是否在“工具箱 + MultiPage”里
Private Function InToolbox() As Boolean
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls("mpTabs")    ' 推荐 MultiPage 名
    InToolbox = (Err.Number = 0)
    Err.Clear
End Function

'（二）安全调用：窗体方法存在则调用（如 CapPage_Init）
Private Function TryCallHostMethod(ByVal host As Object, ByVal methodName As String, ParamArray args()) As Boolean
    On Error Resume Next
    CallByName host, methodName, VbMethod, args
    TryCallHostMethod = (Err.Number = 0)
    Err.Clear
End Function

'（三）安全调用：模块过程存在则运行（支持“模块.过程名”或仅“过程名”）
Private Function RunIfExists(procFullName As String, ParamArray args()) As Boolean
    On Error Resume Next
    Application.Run procFullName, args
    RunIfExists = (Err.Number = 0)
    Err.Clear
End Function

'（四）把“工具箱页面设置”里的值同步到旧窗体（名称：样式_标准化页面设置）
Private Sub CopyPageSetupToOldForm(ByVal oldForm As Object)
    On Error Resume Next
    oldForm.txtTop.text = Me.txtTop.text
    oldForm.txtBottom.text = Me.txtBottom.text
    oldForm.txtLeft.text = Me.txtLeft.text
    oldForm.txtRight.text = Me.txtRight.text

    oldForm.txtTopL.text = Me.txtTopL.text
    oldForm.txtBottomL.text = Me.txtBottomL.text
    oldForm.txtLeftL.text = Me.txtLeftL.text
    oldForm.txtRightL.text = Me.txtRightL.text

    oldForm.txtHeaderLeft.text = Me.txtHeaderLeft.text
    oldForm.txtHeaderRight.text = Me.txtHeaderRight.text
    oldForm.txtLogo.text = Me.txtLogo.text
    oldForm.txtHeaderDist.text = Me.txtHeaderDist.text
    oldForm.txtFooterDist.text = Me.txtFooterDist.text
    On Error GoTo 0
End Sub


' =========================================================
'  一、【样式导入】页（占位：后续把真实实现补进来）
' =========================================================

Private Sub cmdStyleImport_Click()
    Call 一键设置全部样式
End Sub


' =========================================================
'  二、【标题设置】页（占位：后续补真实逻辑）
' =========================================================
Private Sub cmdAutoDetectHeading_Click()
    Call 匹配标题并套用样式_基于配置中心
End Sub
Private Sub cmdMultiLevelMatch_Click()
    Call 标题自动编号
End Sub
Private Sub cmdRemoveManualNumber_Click()
'    Call 去除手工编号_基于配置中心
    Call 去除手工编号_使用进度窗体
End Sub


' =========================
' 三、【页面设置页】：与旧窗体“中交标准化页面设置”联动
' =========================
' —— 页面设置 · 浏览 LOGO ——
Private Sub cmdBrowse_Click()
    '（一）直接让旧窗体执行浏览逻辑；本窗体参数会先同步过去
    样式_标准化页面设置.PS_BrowseLogo Me
End Sub

' —— 页面设置 · 应用到本节 ——
Private Sub cmdApplySection_Click()
    '（二）同步参数 → 旧窗体执行业务
    样式_标准化页面设置.PS_ApplySection Me
End Sub

' —— 页面设置 · 全文应用 ——
Private Sub cmdApplyAll_Click()
    '（三）同步参数 → 旧窗体执行业务（含进度）
    样式_标准化页面设置.PS_ApplyAll Me
End Sub


' =========================================================
'  四、【表格格式化】页
' =========================================================
' —— 全文区域：OK = 直接跑“全文表格格式化（按参数）” ——
Private Sub cmdOK_Click()
    '（一）取左侧三个参数
    Dim nm As String, pt As Single
    nm = Trim(Me.cboFontSize.text)
    If Len(nm) = 0 Then
        MsgBox "请选择中文字号（如“五号”）或输入数字磅值。", vbExclamation
        Exit Sub
    End If

    pt = GetFontSizePt(nm)   ' 你的公共函数
    If pt <= 0 Then
        MsgBox "字号无效：" & nm, vbExclamation
        Exit Sub
    End If

    '（二）直接执行（不再弹出 dlg，不关闭本窗体）
    全文表格格式化_按参数 _
        Me.chkThickOuter.Value, _
        Me.chkFirstRowBold.Value, _
        pt, _
        nm

    '（三）如需提示可加：
    ' MsgBox "全文表格格式化已完成。", vbInformation
End Sub
Private Sub CommandButton23_Click()
    '（一）取左侧三个参数
    Dim nm As String, pt As Single
    nm = Trim(Me.cboFontSize.text)
    If Len(nm) = 0 Then
        MsgBox "请选择中文字号（如“五号”）或输入数字磅值。", vbExclamation
        Exit Sub
    End If

    pt = GetFontSizePt(nm)   ' 你的公共函数
    If pt <= 0 Then
        MsgBox "字号无效：" & nm, vbExclamation
        Exit Sub
    End If

    '（二）直接执行（不再弹出 dlg，不关闭本窗体）
    全文表格格式化_按参数1 _
        Me.chkThickOuter.Value, _
        Me.chkFirstRowBold.Value, _
        pt, _
        nm

    '（三）如需提示可加：
    ' MsgBox "全文表格格式化已完成。", vbInformation
End Sub

Private Sub cmdTF_ApplyAll_Click()
    '（二）全文表格格式化（占位：调用你的总控过程）
    ' 逻辑顺序：
    '  1) 读取“全文区域”的三个选项：外框加粗/首行加粗/字号
    '  2) 调用你的总控过程（例如：全文表格格式化工具），或写进全局配置由总控读取
    '  3) 进度窗体/异常处理
    Dim nm As String, pt As Single
    nm = Trim(cboFontSize.text)
    pt = GetFontSizePt(nm)
    If pt <= 0 Then
        MsgBox "请先选择有效的中文字号或数字磅值。", vbExclamation: Exit Sub
    End If

    ' 【占位调用】你已有的大过程，如无则先占位提示
    If Not RunIfExists("全文表格格式化工具") Then
        MsgBox "【占位】请将全文表格格式化的主过程命名为“全文表格格式化工具”，或在此处改为你的过程名。", vbInformation
    End If
End Sub

' —— 当前表格设置：你已有实现，保留并略做整理 ——
Private Sub cmdApplyCur_Click()
    '（三）当前表格设置 → 调用你已有过程“当前表格格式设置”
    Dim pt As Single
    pt = GetFontSizePt(Me.cboCurFontSize.text)
    If pt <= 0 Then
        MsgBox "请先选择有效的中文字号或数字磅值。", vbExclamation
        Exit Sub
    End If

    Call 当前表格格式设置( _
        Me.chkCurThickOuter.Value, _
        Me.chkCurFirstRowBold.Value, _
        Me.chkCurHeaderRepeat.Value, _
        Me.chkCurAllowBreak.Value, _
        pt)
End Sub

Private Sub cmdTF_Explain_Click()
    '（四）功能说明（占位）
    MsgBox "【占位】展示“全文表格格式化”的说明与注意事项。", vbInformation
End Sub


' =========================================================
'  五、【图表标题】页（事件占位）

'==============================
' （一）按钮事件 → 统一调度
'==============================
Private Sub btnAllMatchStyles_Click()
    调度执行 1   '（1）标题样式匹配
End Sub
Private Sub btnAllAutoNumber_Click()
    调度执行 2   '（2）自动图表编号
End Sub
Private Sub btnAllRemoveManualNo_Click()
    调度执行 3   '（3）去除手工编号
End Sub
Private Sub btnCaptionPreCheck_Click()
    调度执行 4   '（4）标题预检查：表/图根据模式分流
End Sub
Private Sub btnCheckPictures_Click()
    调度执行 5
End Sub
'==============================
' （二）核心调度：根据模式分流到 A/B/C 或 D/E/F
'==============================
Private Sub 调度执行(ByVal 动作编号 As Long)
    Dim modeKey As String
    modeKey = 读取模式Key(Me)   ' 返回 "表" | "图" | ""
    If modeKey = "" Then
        MsgBox "未选择有效模式，请在“模式选择”下拉中选择【表模式/图模式】。", vbExclamation
        Exit Sub
    End If

    Select Case modeKey
        Case "表"
            Select Case 动作编号
                Case 1: Call 表标题样式统一
                Case 2: Call 表格标题自动编号_使用进度窗体
                Case 3: Call 清除表题手工编号_使用进度窗体
                Case 4: Call 自检_表格标题样式一致性
                Case Else: GoTo MAP_ERR
            End Select

        Case "图"
            Select Case 动作编号
                Case 1: Call 图片标题样式统一_带进度
                Case 2: Call 图片标题自动编号_使用进度窗体
                Case 3: Call 清除图片题手工编号_使用进度窗体
                Case 4: Call 自检_图片标题样式一致性
                Case 5: Call 统一图片段落样式_使用进度窗体
                Case Else: GoTo MAP_ERR
            End Select
    End Select
    Exit Sub

MAP_ERR:
    MsgBox "未映射到对应的动作，请检查按钮编号与模式映射。", vbCritical
End Sub

'===============================
' 导入：将控制面板参数覆盖到目标样式（不新建）
' 表模式 → 表格标题；图模式 → 图片标题
'===============================
Private Sub btnCapImport_Click()
    Dim doc As Document: Set doc = ActiveDocument
    Dim modeKey As String, styleName As String
    Dim sty As Style
    Dim fontCN As String, sizeText As String
    Dim sizePt As Single, beforeLines As Single, afterLines As Single
    Dim oneLinePt As Single, boldOn As Boolean
    
    '（一）判定模式（只看首字“表/图”）
    modeKey = 读取模式Key(Me)                 ' 你前面已加过的工具函数
    If modeKey = "" Then
        MsgBox "请先在【模式选择】中选择 表模式/图模式。", vbExclamation
        Exit Sub
    End If
    styleName = IIf(modeKey = "表", "表格标题", "图片标题")
    
    '（二）获取目标样式（不新建）
    On Error Resume Next
    Set sty = doc.Styles(styleName)
    On Error GoTo 0
    If sty Is Nothing Then
        MsgBox "【" & styleName & "】不存在，请先在【样式导入】功能中点击【一键导入】。", vbExclamation
        Exit Sub
    End If
    
    '（三）读取面板参数
    fontCN = NzStr(Me.cboCapFontCN.Value, "黑体")
    sizeText = NzStr(Me.cboCapFontSize.Value, "五号")
    sizePt = 字号到磅值(sizeText, 10.5!)        ' 默认五号≈10.5pt
    boldOn = (Me.chkCapBold.Value = True)
    beforeLines = val(标准化数字文本(Me.txtParaSpaceBeforeLines.text)) ' 允许空=0
    afterLines = val(标准化数字文本(Me.txtParaSpaceAfterLines.text))
    
    '（四）以“字号”作为 1 行的基准 pt 值（常见做法满足你的“行”概念）
    oneLinePt = sizePt
    
    '（五）覆盖必要属性（其余不动）
    With sty.Font
        .NameFarEast = fontCN
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = boldOn
        .Size = sizePt
    End With
    With sty.ParagraphFormat
        .SpaceBefore = beforeLines * oneLinePt
        .SpaceAfter = afterLines * oneLinePt
        ' 不改动：对齐方式、行距、大纲级别、缩进、制表位等
    End With
    
    '（六）让已使用该样式的段落立即刷新一次（不改变文本）
    强制重套样式_简 doc, sty
    
    MsgBox "已更新样式【" & styleName & "】：" & vbCrLf & _
           "中文字体=" & fontCN & "；字号=" & sizeText & "（" & Format(sizePt, "0.0#") & "pt）" & vbCrLf & _
           "加粗=" & IIf(boldOn, "是", "否") & "；段前=" & beforeLines & "行；段后=" & afterLines & "行。", _
           vbInformation
End Sub

' =========================================================
'  六、【图片操作】页面
' =========================================================
Private Sub 全文表格式化_跳过图片_Click()
Call 预处理_标记图片表与普通表
End Sub

Private Sub 设置为图片表_Click()
Call 一键将所选表格标记为图片表
End Sub
Private Sub btnCoverage_Click()
Call 生成封面_绝对坐标
End Sub
Private Sub CommandButton19_Click()
Call 插入单栏图片表_图片控件版
End Sub
Private Sub CommandButton22_Click()
Call 插入双栏图片表_图片控件版_双栏
End Sub


' =========================================================
'  七、窗体级别的退出/取消
' =========================================================
Private Sub cmdCancel_Click()
    Canceled = True
    Unload Me
End Sub


'==============================
' （三）工具：读取模式键、运行宏
'==============================
' 读取下拉值，返回首字 "表"/"图"（其余返回空）
Private Function 读取模式Key(ByVal host As Object) As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(host.cboModeSelect.Value))
    On Error GoTo 0
    If Len(v) = 0 Then Exit Function
    v = Replace$(v, "模式", "")         ' 去掉“模式”二字（如：表模式/图模式）
    v = Left$(v, 1)                     ' 只取首字
    If v = "表" Or v = "图" Then 读取模式Key = v
End Function

' 按名称运行 Public Sub（在标准模块中），并给出友好错误提示
Private Sub 安全运行宏(ByVal macroName As String)
    On Error GoTo EH
    Application.Run macroName
    Exit Sub
EH:
    MsgBox "未找到可执行过程：" & macroName & vbCrLf & _
           "请确认该过程存在于“标准模块”且为 Public Sub。", vbExclamation
End Sub
'——（工具）空/Null 变成默认字符串
Private Function NzStr(ByVal v As Variant, ByVal def As String) As String
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then NzStr = def Else NzStr = CStr(v)
End Function

'——（工具）把“10.5 / １０．５ / 10.5pt”这类文本标准化为可 Val 的半角数字
Private Function 标准化数字文本(ByVal s As String) As String
    s = Trim$(s)
    s = Replace$(s, "．", "."): s = Replace$(s, "。", ".")
    s = Replace$(s, "，", "."): s = Replace$(s, "、", ".")
    s = Replace$(s, "－", "-"): s = Replace$(s, "—", "-")
    s = Replace$(s, "＋", "+")
    s = Replace$(s, "pt", "", , , vbTextCompare)
    s = Replace$(s, "ＰＴ", "", , , vbTextCompare)
    标准化数字文本 = s
End Function

'——（工具）中文字号 → pt；若未识别则尝试数值，最终回退默认值
Private Function 字号到磅值(ByVal s As String, ByVal defPt As Single) As Single
    Dim key As String: key = Trim$(s)
    Select Case key
        Case "初号": 字号到磅值 = 42#
        Case "小初": 字号到磅值 = 36#
        Case "一号": 字号到磅值 = 26#
        Case "小一": 字号到磅值 = 24#
        Case "二号": 字号到磅值 = 22#
        Case "小二": 字号到磅值 = 18#
        Case "三号": 字号到磅值 = 16#
        Case "小三": 字号到磅值 = 15#
        Case "四号": 字号到磅值 = 14#
        Case "小四": 字号到磅值 = 12#
        Case "五号": 字号到磅值 = 10.5
        Case "小五": 字号到磅值 = 9#
        Case "六号": 字号到磅值 = 7.5
        Case "小六": 字号到磅值 = 6.5
        Case "七号": 字号到磅值 = 5.5
        Case "八号": 字号到磅值 = 5#
        Case Else
            ' 允许直接输入 10.5 / 11 等
            key = 标准化数字文本(key)
            If IsNumeric(key) Then
                字号到磅值 = CSng(key)
            Else
                字号到磅值 = defPt
            End If
    End Select
End Function

'——（工具）把文档中已应用某样式的段落“重套一次”以立刻生效（无显式循环）
Private Sub 强制重套样式_简(ByVal doc As Document, ByVal 目标样式 As Style)
    With doc.content.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .text = ""
        .replacement.text = ""
        .Style = 目标样式
        .replacement.Style = 目标样式
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

