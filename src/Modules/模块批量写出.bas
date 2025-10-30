Attribute VB_Name = "模块批量写出"
Option Explicit
'======================================================================
'  Word VBA | 导出当前工程到 GitHub 友好结构（覆盖式 + README 只更新自动区）
'  Entry : 导出_当前工程_白名单 / 导出_当前工程_全部
'======================================================================

'============================（一）导出配置============================
'【要求】导出根目录（自动创建）
Private Const BACKUP_ROOT As String = "E:\BaiduSyncdisk\Word-ToolBox-VBA"
'【要求】导出前是否清空 src 目录（True=清空，避免旧文件遗留）
Private Const CLEAR_SRC_BEFORE_EXPORT As Boolean = True

' README 自动区块的标记（只替换这段，保留你手写的其它内容）
Private Const MARK_BEGIN As String = "<!-- AUTO:EXPORT-BLOCK:BEGIN -->"
Private Const MARK_END   As String = "<!-- AUTO:EXPORT-BLOCK:END -->"

'（一）精确名称白名单（大小写不敏感；你后续只需在此处追加）
Private Function NamesToExport() As Variant
    NamesToExport = Array( _
        "frmDocObjectInspector", "ProgressForm", "标准化样式工具箱", _
        "全文跨行断页功能", "样式_标准化页面设置", "样式_字体样式一键导入", "自检报告", _
        "MOD_窗体初始化中心", "MOD_窗体唤醒", "MOD_公共工具函数", "MOD_样式中心", _
        "编号_1_导入各级大纲样式", "编号_2_自动匹配各级标题", "编号_3_多级模板定义和匹配", _
        "编号_4_自动删除手工编号", "编号_5_自检工具", "编号_6_带进度的去除编号", _
        "表格_首行加粗", _
        "表格标题_1_统一表标题样式", "表格标题_1_统一_表标题样式", _
        "表格标题_2_自动表标题编号", "表格标题_3_去除人工编号", "表格标题_4_人工检查工具", _
        "表格标题_5_预检查", "表格标题_6_极简自检", _
        "全文_表格_标题加粗", "全文_表格_跨行断页", "全文_表格_统一字体样式", "全文_表格_完全格式化", _
        "全文_全选当前表", "全文_样式_清除未使用样式", _
        "图片标题_1_样式匹配", "图片标题_2_自动编号", "图片标题_3_去除人工编号", _
        "样式_标准化表格样式", "样式_施工组织样式一键导入", "样式_四级标题_表格标题", "样式_正文_施工组织设计" _
    )
End Function

'（二）通配符白名单——当前关闭（保持空数组；以后需要再加）
Private Function PatternsToExport() As Variant
    PatternsToExport = Array()
End Function

'============================（二）公开入口============================
'（一）只导出当前工程：按白名单
Public Sub 导出_当前工程_白名单()
    ExportActiveProject False
End Sub

'（二）只导出当前工程：全部组件
Public Sub 导出_当前工程_全部()
    ' 弹出输入面板获取Commit信息
    Dim commitMsg As String
    commitMsg = InputBox( _
        Prompt:="请输入本次导出的Commit信息（将显示在README中）：" & vbCrLf & "例如：修复表格标题编号逻辑", _
        Title:="输入Commit信息", _
        Default:="") ' 默认空值
    
    ' 处理用户取消/未输入的情况
    If commitMsg = "" Then
        MsgBox "未输入Commit信息，已取消导出操作。", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 调用核心导出过程，并传递commit信息
    ExportActiveProject True, commitMsg

End Sub

'============================（三）核心过程============================
Private Sub ExportActiveProject(ByVal exportAll As Boolean, Optional ByVal commitMsg As String = "")
    On Error GoTo FAIL

    '（一）拿当前工程
    Dim proj As Object
    On Error Resume Next
    Set proj = Application.VBE.ActiveVBProject
    On Error GoTo 0
    If proj Is Nothing Then
        MsgBox "未找到活动 VBA 工程。请在工程窗口选中一个工程后再试。", vbExclamation, "导出"
        Exit Sub
    End If
    If Not CanAccessProject(proj) Then
        MsgBox Join(Array( _
            "当前工程被保护或未启用“信任对VBA工程对象模型的访问”。", _
            "请到：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置，勾选该项。" _
        ), vbCrLf), vbCritical, "无法访问工程"
        Exit Sub
    End If

    '（二）准备 Git 友好目录
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim root As String: root = EnsureFolder(BACKUP_ROOT, fso)
    Dim dirSrc As String: dirSrc = EnsureFolder(root & "\src", fso)
    Call EnsureFolder(dirSrc & "\Modules", fso)
    Call EnsureFolder(dirSrc & "\Classes", fso)
    Call EnsureFolder(dirSrc & "\Forms", fso)
    Call EnsureFolder(dirSrc & "\Documents", fso)
    Call EnsureFolder(root & "\manifest", fso)
    Call EnsureFolder(root & "\templates\Normal", fso)

    '（三）可选进度窗
    Dim pf As Object
    On Error Resume Next
    Set pf = VBA.UserForms.Add("ProgressForm")
    On Error GoTo 0
    If Not pf Is Nothing Then
        pf.caption = "导出（当前工程）：" & proj.name
        pf.Show vbModeless
        DoEvents
    End If

    '（四）统计数量
    Dim total As Long: total = CountMatchesInProject(proj, exportAll)
    If total = 0 Then
        StatusPulse pf, "未匹配到需要导出的组件。"
        GoTo WRITE_META
    End If

    '（五）清空旧文件（仅 src/* 目录）
    If CLEAR_SRC_BEFORE_EXPORT Then
        PurgeFolder dirSrc & "\Modules"
        PurgeFolder dirSrc & "\Classes"
        PurgeFolder dirSrc & "\Forms"
        PurgeFolder dirSrc & "\Documents"
        StatusPulse pf, "已清空旧的导出文件（src/*）。"
    End If

    '（六）导出组件（覆盖）
    Dim done As Long: done = 0
    Dim comp As Object, log As String
    For Each comp In proj.VBComponents
        If exportAll Or IsTargetName(comp.name) Then
            Dim subDir As String, ext As String, dst As String, base As String
            subDir = SubFolderByType(comp.Type)
            ext = GuessExtByType(comp.Type)
            base = root & "\src\" & subDir & "\" & SafeFile(comp.name)
            dst = base & ext

            ' 覆盖旧文件（含 .frx）
            If fso.FileExists(dst) Then On Error Resume Next: fso.DeleteFile dst, True: On Error GoTo 0
            If ext = ".frm" Then
                If fso.FileExists(base & ".frx") Then On Error Resume Next: fso.DeleteFile base & ".frx", True: On Error GoTo 0
            End If

            On Error Resume Next
            comp.Export dst
            If Err.Number = 0 Then
                log = log & "√ " & comp.name & "  →  " & dst & vbCrLf
            Else
                log = log & "× " & comp.name & "  →  " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0

            done = done + 1
            UpdateBar pf, done, total, "导出（" & done & "/" & total & "）： " & comp.name
            If Not pf Is Nothing Then If pf.stopFlag Then Exit For
        End If
    Next

    '（七）同时备份 Normal.dotm（覆盖）
    CopyNormalTemplate root & "\templates\Normal\normal.dotm"

WRITE_META:
    '（八）写引用清单 / 导出日志 / Git 文件
    WriteReferences proj, root & "\manifest\references.txt"
    WriteTextFile root & "\manifest\export_log.txt", log
    WriteGitignore root
    WriteReadme root, proj.name, commitMsg    ' ← 新增传递commit信息

    StatusPulse pf, "完成。共匹配到 " & CStr(total) & " 个组件，已处理 " & CStr(done) & " 个。"
    If Not pf Is Nothing Then Unload pf
    Application.StatusBar = False
    MsgBox "导出完成，根目录：" & root, vbInformation, "导出"
    Exit Sub

FAIL:
    If Not pf Is Nothing Then Unload pf
    Application.StatusBar = False
    MsgBox "导出过程中发生错误：" & Err.Description, vbCritical, "导出失败"
End Sub

'============================（四）判定/统计/映射============================
Private Function CanAccessProject(ByVal proj As Object) As Boolean
    On Error Resume Next
    Dim c As Long: c = proj.VBComponents.Count
    CanAccessProject = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function IsTargetName(ByVal compName As String) As Boolean
    Dim nm As Variant
    For Each nm In NamesToExport
        If UCase$(CStr(nm)) = UCase$(compName) Then IsTargetName = True: Exit Function
    Next
    For Each nm In PatternsToExport
        If compName Like CStr(nm) Then IsTargetName = True: Exit Function
    Next
End Function

Private Function CountMatchesInProject(ByVal proj As Object, ByVal exportAll As Boolean) As Long
    Dim comp As Object, N As Long
    If Not CanAccessProject(proj) Then Exit Function
    For Each comp In proj.VBComponents
        If exportAll Or IsTargetName(comp.name) Then N = N + 1
    Next
    CountMatchesInProject = N
End Function

Private Function SubFolderByType(ByVal vbCompType As Long) As String
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2
    Const vbext_ct_MSForm As Long = 3
    Const vbext_ct_Document As Long = 100
    Select Case vbCompType
        Case vbext_ct_StdModule:   SubFolderByType = "Modules"
        Case vbext_ct_ClassModule: SubFolderByType = "Classes"
        Case vbext_ct_MSForm:      SubFolderByType = "Forms"
        Case vbext_ct_Document:    SubFolderByType = "Documents"
        Case Else:                 SubFolderByType = "Modules"
    End Select
End Function

Private Function GuessExtByType(ByVal vbCompType As Long) As String
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2
    Const vbext_ct_MSForm As Long = 3
    Const vbext_ct_Document As Long = 100
    Select Case vbCompType
        Case vbext_ct_StdModule:   GuessExtByType = ".bas"
        Case vbext_ct_ClassModule: GuessExtByType = ".cls"
        Case vbext_ct_MSForm:      GuessExtByType = ".frm"   ' 会伴随 .frx
        Case vbext_ct_Document:    GuessExtByType = ".cls"
        Case Else:                 GuessExtByType = ".bas"
    End Select
End Function

'============================（五）IO / Git / 状态工具============================
' 确保目录存在
' 确保目录存在（递归创建，支持多级/UNC/尾反斜杠）
Private Function EnsureFolder(ByVal path As String, ByVal fso As Object) As String
    Dim p As String: p = path
    ' 去掉末尾 "\"（避免把根目录误删到空串以外）
    If Len(p) > 3 And Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If p = "" Then EnsureFolder = "": Exit Function

    ' 已存在直接返回
    If fso.FolderExists(p) Then
        EnsureFolder = p
        Exit Function
    End If

    ' 先确保上级目录存在
    Dim cut As Long: cut = InStrRev(p, "\")
    If cut > 0 Then
        Dim parent As String: parent = Left$(p, cut - 1)
        ' 对于盘符根（如 "C:"）和 UNC 根（如 "\\server\share"）做保护
        If parent <> "" And Not fso.FolderExists(parent) Then
            ' 若是 UNC，确保到 \\server\share 为止
            If Left$(p, 2) = "\\" Then
                Dim first As Long, second As Long
                first = InStr(3, p, "\")                   ' 第三个字符之后找第一个 "\"
                If first > 0 Then second = InStr(first + 1, p, "\") ' 共享名后的 "\"
                If second > 0 And Len(parent) >= second - 1 Then
                    ' parent 仍然可能深一层，递归即可
                End If
            End If
            Call EnsureFolder(parent, fso)
        End If
    End If

    ' 创建当前目录（若并发/权限失败用 MkDir 兜底）
    On Error Resume Next
    fso.CreateFolder p
    If Err.Number <> 0 Then
        Err.Clear
        MkDir p
    End If
    On Error GoTo 0

    EnsureFolder = p
End Function


' 清空目录中文件（不删子目录）
Private Sub PurgeFolder(ByVal folderPath As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        Dim f As Object
        For Each f In fso.GetFolder(folderPath).Files
            f.Delete True
        Next
    End If
    On Error GoTo 0
End Sub

' 写文本（UTF-8 BOM，覆盖）
Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2: .Charset = "utf-8": .Open
        .WriteText content
        .SaveToFile filePath, 2   ' 2=覆盖
        .Close
    End With
End Sub

' 读文本（UTF-8 优先，失败退回 ANSI）
Private Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo FAIL
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2: .Charset = "utf-8": .Open
        .LoadFromFile filePath
        ReadTextFile = .ReadText
        .Close
    End With
    Exit Function
FAIL:
    On Error GoTo 0
    If Len(Dir$(filePath)) > 0 Then
        Dim f As Integer: f = FreeFile
        Open filePath For Input As #f
        ReadTextFile = Input$(LOF(f), f)
        Close #f
    Else
        ReadTextFile = ""
    End If
End Function

' 写引用清单
Private Sub WriteReferences(ByVal proj As Object, ByVal outFile As String)
    On Error Resume Next
    Dim ref As Object, lines As Variant, s As String
    s = "Project: " & proj.name & vbCrLf & _
        "Exported: " & Format(Now, "yyyy-mm-dd HH:NN:SS") & vbCrLf & _
        String(40, "-") & vbCrLf
    For Each ref In proj.References
        s = s & ref.name & " | " & ref.GUID & " | v" & ref.Major & "." & ref.Minor & _
            " | " & ref.FullPath & vbCrLf
    Next
    On Error GoTo 0
    WriteTextFile outFile, s
End Sub

' 写 .gitignore（覆盖）——用 Join 降低行继续符
Private Sub WriteGitignore(ByVal root As String)
    Dim lines As Variant
    lines = Array( _
        "# Office & VBA", _
        "~$*", _
        "*.tmp", _
        "*.lock", _
        "Thumbs.db", _
        ".DS_Store", _
        "" _
    )
    WriteTextFile root & "\.gitignore", Join(lines, vbCrLf)
End Sub

' 写 README（仅首次创建模板；以后仅更新自动区块）
Private Sub WriteReadme(ByVal root As String, ByVal projName As String, ByVal commitMsg As String)
    Dim path As String: path = root & "\README.md"
    Dim content As String, exists As Boolean
    exists = (Len(Dir$(path)) > 0)
    Dim autoBlock As String: autoBlock = BuildReadmeAutoBlock(projName, commitMsg)

    If Not exists Then
        ' 第一次生成：完整模板 + 自动区块
        WriteTextFile path, BuildReadmeTemplate(projName, autoBlock)
    Else
        ' 后续运行：仅替换自动区块，保留用户手工内容
        content = ReadTextFile(path)
        If InStr(1, content, MARK_BEGIN, vbTextCompare) > 0 And _
           InStr(1, content, MARK_END, vbTextCompare) > 0 Then
            content = ReplaceAutoBlock(content, MARK_BEGIN, MARK_END, autoBlock)
        Else
            content = content & vbCrLf & vbCrLf & autoBlock
        End If
        WriteTextFile path, content
    End If
End Sub
' 生成 README 模板（少行继续符）
Private Function BuildReadmeTemplate(ByVal projName As String, ByVal autoBlock As String) As String
    Dim lines As Variant
    lines = Array( _
        "# " & projName & "（Word 插件 / VBA 项目）", "", "本仓库存放 **当前 Word VBA 工程** 的源码导出，便于版本管理与协作。", _
        "导出脚本会覆盖旧版本，并将 Normal.dotm 一并备份。", "", "## 目录结构", "- `src/Modules/`：标准模块（.bas）", "- `src/Classes/`：类模块（.cls）", _
        "- `src/Forms/`：窗体（.frm + .frx）", "- `src/Documents/`：文档模块（.cls，ThisDocument等）", "- `templates/Normal/normal.dotm`：Normal 模板备份", _
        "- `manifest/references.txt`：工程引用清单", "- `manifest/export_log.txt`：导出日志", "", "## 使用方法", "1. 在 Word 按 `Alt+F11` 打开 VBA 编辑器。", _
        "2. 运行宏：`导出_当前工程_白名单`（或 `导出_当前工程_全部`）。", "3. 导出根目录：`" & BACKUP_ROOT & "`。", "4. 首次使用：请在 *文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置* 勾选“**信任对 VBA 工程对象模型的访问**”。", _
        "", "## 首次推送到 GitHub（示例）", "```bash", "cd """ & BACKUP_ROOT & """", "git init", "git add .", "git commit -m ""init: 导出当前工程源码""", _
        "git branch -M main", "git remote add origin https://github.com/<your-account>/<repo>.git", "git push -u origin main", _
        "```", _
        "", _
        "## 约定", _
        "- 模块/过程注释采用（**一**/**二**/…）编号，便于审阅。", _
        "- 命名保持前缀语义（如 `MOD_`、`表格标题_`、`全文_`）。", _
        "- 逐步迁移到 VSTO/VB.NET 时，公共逻辑保持在独立模块中，减少耦合。", _
        "", _
        autoBlock _
    )
    BuildReadmeTemplate = Join(lines, vbCrLf)
End Function

' 生成 README 的“自动区块”（只此段会被覆盖）
Private Function BuildReadmeAutoBlock(ByVal projName As String, ByVal commitMsg As String) As String
    Dim lines As Variant
    lines = Array( _
        MARK_BEGIN, _
        "### 导出信息（自动生成）", _
        "- 工程名： " & projName, _
        "- 导出时间： " & Format(Now, "yyyy-mm-dd HH:NN:SS"), _
        "- 根目录： " & BACKUP_ROOT, _
        "- 本次更新： " & commitMsg, _
        MARK_END _
    )
    BuildReadmeAutoBlock = Join(lines, vbCrLf)
End Function

' 用新的自动区块替换旧区块
Private Function ReplaceAutoBlock(ByVal content As String, _
                                  ByVal tagBegin As String, _
                                  ByVal tagEnd As String, _
                                  ByVal newBlock As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, content, tagBegin, vbTextCompare)
    p2 = InStr(p1 + Len(tagBegin), content, tagEnd, vbTextCompare)
    If p1 > 0 And p2 > p1 Then
        ReplaceAutoBlock = Left$(content, p1 - 1) & newBlock & mid$(content, p2 + Len(tagEnd))
    Else
        ReplaceAutoBlock = content & vbCrLf & vbCrLf & newBlock
    End If
End Function

' 复制 Normal.dotm（覆盖）
Private Sub CopyNormalTemplate(ByVal dstFullPath As String)
    On Error Resume Next
    Dim src As String: src = Application.NormalTemplate.FullName
    If Len(Dir$(src)) > 0 Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        EnsureFolder Left$(dstFullPath, InStrRev(dstFullPath, "\") - 1), fso
        fso.CopyFile src, dstFullPath, True
        StatusPulse Nothing, "已备份 Normal.dotm → " & dstFullPath
    End If
    On Error GoTo 0
End Sub

' 状态输出 / 进度
Private Sub StatusPulse(ByVal pf As Object, ByVal msg As String)
    On Error Resume Next
    If Not pf Is Nothing Then
        pf.TextBoxStatus.text = pf.TextBoxStatus.text & vbCrLf & msg
        pf.TextBoxStatus.SelStart = Len(pf.TextBoxStatus.text)
        pf.TextBoxStatus.SelLength = 0
        pf.Repaint
    End If
    Application.StatusBar = msg
    DoEvents
    On Error GoTo 0
End Sub

Private Sub UpdateBar(ByVal pf As Object, ByVal cur As Long, ByVal total As Long, ByVal msg As String)
    On Error Resume Next
    Dim w As Integer: If total > 0 Then w = CInt(200 * (CDbl(cur) / total))
    If Not pf Is Nothing Then pf.UpdateProgressBar w, msg Else Application.StatusBar = msg
    On Error GoTo 0
End Sub

' 文件名净化
Private Function SafeFile(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace$(s, CStr(bad), "_")
    Next
    SafeFile = s
End Function

