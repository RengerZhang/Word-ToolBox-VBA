Attribute VB_Name = "MOD_窗体初始化中心"
Option Explicit
' =========================================================
'  MOD_格式初始化中心
'  作用：统一管理【工具箱窗体】各子页面的初始化（懒加载）
'  约定：
'   1）MultiPage 控件名推荐：mpTabs（若不是，也能自动识别）
'   2）各页 Page.Name：
'        - pgPageSetup     页面设置
'        - pgCaption       图表标题
'        - pgTableFormat   表格格式化
'        - pgTitle         标题设置（占位）
'        - pgStyleImport   样式导入（占位）
'   3）每页仅在首次或 force=True 时执行（用 Page.Tag="inited" 标记）
' =========================================================


'（一）对外入口：按“页面名称”初始化（懒加载）
Public Sub Init_ByPageName(ByVal host As Object, ByVal pageName As String, Optional ByVal force As Boolean = False)
    Select Case LCase$(pageName)
        Case "pgpagesetup":    Init_PageSetup host, force
        Case "pgcaption":      Init_Caption host, force
        Case "pgtableformat":  Init_TableFormat host, force
        Case "pgtitle":        Init_Title host, force            ' 占位
        Case "pgstyleimport":  Init_StyleImport host, force      ' 占位
        Case Else
            ' 未知页名：不处理
    End Select
End Sub


'（二）对外入口：初始化“当前选中的页面”（用于窗体 Initialize / MultiPage_Change）
Public Sub Init_CurrentPage(ByVal host As Object, Optional ByVal force As Boolean = False)
    Dim mp As Object: Set mp = FindMultiPage(host)
    If mp Is Nothing Then Exit Sub
    Dim curName As String
    curName = mp.Pages(mp.Value).name
    Init_ByPageName host, curName, force
End Sub


'（三）可选入口：一次性初始化全部页面（需要“恢复默认全部页面”时调用）
Public Sub Init_All(ByVal host As Object, Optional ByVal force As Boolean = False)
    Init_PageSetup host, force
    Init_Caption host, force
    Init_TableFormat host, force
    Init_Title host, force
    Init_StyleImport host, force
End Sub


' =========================================================
'  具体页面初始化（每页各一段，便于单步调试）
' =========================================================

'（四）页面设置页：这里“写死”默认值（不再调用 ps_Init）
Public Sub Init_PageSetup(ByVal host As Object, Optional ByVal force As Boolean = False)
      If Not force Then
        Dim pg As Object: Set pg = GetPage(host, "pgPageSetup")
        If Not pg Is Nothing Then
            If CStr(pg.tag) = "inited" Then Exit Sub
        End If
    End If
    On Error Resume Next     '（一）控件缺失时自动跳过，避免报错

    ' —— 1. 竖版边距（cm）——
    host.txtTop.text = "2.5"
    host.txtBottom.text = "2.5"
    host.txtLeft.text = "3"
    host.txtRight.text = "3"

    ' —— 2. 横版边距（cm）——
    host.txtTopL.text = "3"
    host.txtBottomL.text = "3"
    host.txtLeftL.text = "2.5"
    host.txtRightL.text = "2.5"

    ' —— 3. 页眉文字 ——（左侧两行 / 右侧一行）
    host.txtHeaderLeft.text = "闵行区华漕镇 MHP0-1403单元" & vbCrLf & "73-04地块征收(动迁)安置住房项目"
    host.txtHeaderRight.text = "施工组织设计"

    ' —— 4. 页眉/页脚距离（cm）——
    host.txtHeaderDist.text = "1.5"
    host.txtFooterDist.text = "1.75"

    ' —— 5. Logo 路径（默认留空，不强填）——
    If Len(host.txtLogo.text) = 0 Then host.txtLogo.text = "C:\Users\Tony Zhang\Desktop\logo.png"

    On Error GoTo 0

    MarkInited host, "pgPageSetup"
End Sub

'（五）图表标题页：初始化“模式选择/中文字体/字号”下拉（无判定，强制执行）
Public Sub Init_Caption(ByVal host As Object, Optional ByVal force As Boolean = False)
    On Error Resume Next   '（防止单个控件缺失时报错中断）

    '（一）模式选择：仅两项，禁止手输（DropDownList + MatchRequired）
    With host.cboModeSelect
        SetAsDropDownListSafe host, .name
        .Clear
        .AddItem "表模式"
        .AddItem "图模式"
        .ListIndex = 0          ' 默认：表模式
        .ListRows = 6
    End With

    '（二）中文字体下拉：常用集合，禁止手输
    With host.cboCapFontCN
        SetAsDropDownListSafe host, .name
        .Clear
        AddChineseFonts .Object ' 你已有/上文提供的填充函数
        .Value = "黑体"
        .ListRows = 12
    End With

    '（三）字号下拉：中文字号 + 常用磅值，禁止手输
    With host.cboCapFontSize
       .Style = fmStyleDropDownList
        .MatchRequired = True
        .Clear
        AddChineseFontSizes .Object ' 你已有的字号填充函数
        .Value = "五号"
        .ListRows = 12
    End With

    On Error GoTo 0
End Sub




'（六）表格格式化页：当前需求的默认值/下拉（中文字号 + 常用磅值）
Public Sub Init_TableFormat(ByVal host As Object, Optional ByVal force As Boolean = False)
      If Not force Then
        Dim pg As Object: Set pg = GetPage(host, "pgPageSetup")
        If Not pg Is Nothing Then
            If CStr(pg.tag) = "inited" Then Exit Sub
        End If
    End If
    
    ' —— 1) 全文设置：字号下拉 ——
    With host.cboFontSize
        SetAsDropDownListSafe host, .name
        .Clear
        AddChineseFontSizes .Object
        .Value = "五号"
        .ListRows = 12
    End With

    ' —— 2) 当前表格：字号下拉 ——
    With host.cboCurFontSize
        SetAsDropDownListSafe host, .name
        .Clear
        AddChineseFontSizes .Object
        .Value = "五号"
        .ListRows = 12
    End With

    ' —— 3) 全文格式化：默认开关 ——
    host.chkThickOuter.Value = True        ' 外框加粗：开
    host.chkFirstRowBold.Value = True      ' 首行加粗：开

    ' —— 4) 当前表格设置：默认开关 ——
    host.chkCurThickOuter.Value = True     ' 外框加粗：开
    host.chkCurFirstRowBold.Value = False  ' 首行加粗：关
    host.chkCurHeaderRepeat.Value = True   ' 首行重复：开
    host.chkCurAllowBreak.Value = False    ' 跨行断页：关

    MarkInited host, "pgTableFormat"
End Sub


'（七）标题设置页（占位：不报错；后续你把具体初始化补到这里）
Public Sub Init_Title(ByVal host As Object, Optional ByVal force As Boolean = False)
      If Not force Then
        Dim pg As Object: Set pg = GetPage(host, "pgPageSetup")
        If Not pg Is Nothing Then
            If CStr(pg.tag) = "inited" Then Exit Sub
        End If
    End If
    ' TODO：标题页控件初始化（下拉/默认值/联动）
    MarkInited host, "pgTitle"
End Sub


'（八）样式导入页（占位：不报错；后续你把具体初始化补到这里）
Public Sub Init_StyleImport(ByVal host As Object, Optional ByVal force As Boolean = False)
    If Not PageNeedsInit(host, "pgStyleImport", force) Then Exit Sub
    ' TODO：样式导入页初始化（默认路径、列表、按钮状态等）
    MarkInited host, "pgStyleImport"
End Sub


' =========================================================
'  私有工具（请保留）：页面存在/初始化标记/安全下拉/通用下拉/安全调用
' =========================================================

'（九）查找 MultiPage（优先找名为 mpTabs；否则取第一个 MultiPage）
Private Function FindMultiPage(ByVal host As Object) As Object
    On Error Resume Next
    Set FindMultiPage = host.Controls("mpTabs")
    If FindMultiPage Is Nothing Then
        Dim ctl As Object
        For Each ctl In host.Controls
            If TypeName(ctl) = "MultiPage" Then
                Set FindMultiPage = ctl
                Exit For
            End If
        Next
    End If
End Function

'（十）取 Page（不存在返回 Nothing）
Private Function GetPage(ByVal host As Object, ByVal pageName As String) As Object
    Dim mp As Object: Set mp = FindMultiPage(host)
    If mp Is Nothing Then Exit Function
    On Error Resume Next
    Set GetPage = mp.Pages(pageName)
End Function

'（十一）是否需要初始化（不存在即不初始化；force=True 强制）
Private Function PageNeedsInit(ByVal host As Object, ByVal pageName As String, ByVal force As Boolean) As Boolean
    Dim pg As Object: Set pg = GetPage(host, pageName)
    If pg Is Nothing Then Exit Function
    If force Then
        PageNeedsInit = True
    Else
        PageNeedsInit = (CStr(pg.tag) <> "inited")
    End If
End Function

'（十二）标记为已初始化
Private Sub MarkInited(ByVal host As Object, ByVal pageName As String)
    Dim pg As Object: Set pg = GetPage(host, pageName)
    If Not pg Is Nothing Then pg.tag = "inited"
End Sub

'（十三）把 ComboBox 设为“不可手输 + 必须为列表项”（安全设置）
Private Sub SetAsDropDownListSafe(ByVal host As Object, ByVal comboName As String)
    On Error Resume Next
    With host.Controls(comboName)
        .Style = fmStyleDropDownList
        .MatchRequired = True
    End With
End Sub

'（十四）通用：填充“中文字号 + 常用磅值”（显示项；转磅值请用你公共函数 GetFontSizePt）
Private Sub AddChineseFontSizes(ByVal cbo As Object)
    With cbo
        .AddItem "初号": .AddItem "小初": .AddItem "一号": .AddItem "小一"
        .AddItem "二号": .AddItem "小二": .AddItem "三号": .AddItem "小三"
        .AddItem "四号": .AddItem "小四": .AddItem "五号": .AddItem "小五"
        .AddItem "六号": .AddItem "小六"
        .AddItem "8": .AddItem "9": .AddItem "10": .AddItem "12"
        .AddItem "14": .AddItem "16": .AddItem "18"
    End With
End Sub
'（十四-B）通用：填充常用中文/西文字体（可按需扩展）
Private Sub AddChineseFonts(ByVal cbo As Object)
    With cbo
        .AddItem "黑体"
        .AddItem "宋体"
        .AddItem "仿宋"
        .AddItem "楷体"
        .AddItem "微软雅黑"
        .AddItem "Times New Roman"  ' 字母/数字常用
    End With
End Sub


'（十五）安全调用：窗体方法存在则调用（如 CapPage_Init）
Private Function TryCallHostMethod(ByVal host As Object, ByVal methodName As String, ByVal force As Boolean) As Boolean
    On Error Resume Next
    CallByName host, methodName, VbMethod, force
    TryCallHostMethod = (Err.Number = 0)
    Err.Clear
End Function

'（十六）安全调用：模块过程存在则运行（支持“模块.过程名”或仅“过程名”）
Private Function RunIfExists(procFullName As String, ByVal host As Object, ByVal force As Boolean) As Boolean
    On Error Resume Next
    Application.Run procFullName, host, force
    RunIfExists = (Err.Number = 0)
    Err.Clear
End Function

