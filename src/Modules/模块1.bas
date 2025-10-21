Attribute VB_Name = "模块1"
Option Explicit

'==========================================================
' 一键打包 Normal.dotm 中的指定组件到独立 .dotm
' 使用方法：
' 1）修改 ①处的“要打包的组件清单”
' 2）运行 Pack_Normal_Modules_ToDotm
' 3）在 ②处修改输出路径（默认到桌面）
'==========================================================
Sub Pack_Normal_Modules_ToDotm()
    On Error GoTo ERRH

    '（一）① 定义要打包的组件清单（名称必须与 VBE 中一致）
    Dim modNames As Variant, clsNames As Variant, frmNames As Variant
    modNames = Array( _
        "modCommon", _               ' 示例：公共函数
        "modTitleMatch", _           ' 示例：标题匹配
        "modAutoNumber", _           ' 示例：自动编号
        "modRemoveManualNo" _        ' 示例：删除手工编号
    )
    clsNames = Array()               ' 如有类模块，填入名称
    frmNames = Array( _
        "ProgressForm", _            ' 示例：进度窗体
        "PageSettings"               ' 示例：页面窗体
    )

    '（二）② 输出路径与文件名
    Dim outDir As String, outName As String, outFull As String, stamp As String
    outDir = Environ$("USERPROFILE") & "\Desktop\打包输出\"    ' 默认桌面/打包输出/
    EnsureFolderExists outDir
    stamp = Format(Now, "yyyymmdd_HHNNSS")
    outName = "Word格式工具箱_" & stamp & ".dotm"
    outFull = outDir & outName

    '（三）导出选定组件到临时目录
    Dim tmpRoot As String
    tmpRoot = Environ$("TEMP") & "\vba_pack_" & stamp
    EnsureFolderExists tmpRoot
    Dim dirMod$, dirCls$, dirFrm$
    dirMod = tmpRoot & "\Modules": EnsureFolderExists dirMod
    dirCls = tmpRoot & "\Classes": EnsureFolderExists dirCls
    dirFrm = tmpRoot & "\Forms":   EnsureFolderExists dirFrm

    Application.NormalTemplate.Save
    Dim vbproj As Object: Set vbproj = Application.NormalTemplate.VBProject
    Call ExportComponents(vbproj, modNames, dirMod, 1) ' 1=标准模块
    Call ExportComponents(vbproj, clsNames, dirCls, 2) ' 2=类模块
    Call ExportComponents(vbproj, frmNames, dirFrm, 3) ' 3=窗体

    '（四）新建文档 → 导入组件 → 另存为 .dotm
    Dim doc As Document: Set doc = Documents.Add
    Dim target As Object: Set target = doc.VBProject
    Call ImportComponents(target, dirMod, "*.bas")
    Call ImportComponents(target, dirCls, "*.cls")
    Call ImportComponents(target, dirFrm, "*.frm")

    '（五）写入“引导模块”（创建按钮 / 自安装 / 卸载）
    Call InjectBootstrapModule(target)

    '（六）保存为 .dotm 模板并关闭
    doc.SaveAs2 FileName:=outFull, FileFormat:=wdFormatXMLTemplateMacroEnabled
    doc.Close SaveChanges:=True

    MsgBox "打包完成：" & vbCrLf & outFull, vbInformation
    Exit Sub

ERRH:
    MsgBox "打包失败（" & Err.Number & "）： " & Err.Description, vbCritical
End Sub

'------------------------------------------
' 导出：按名称数组导出指定类型组件
' compType: 1=标准模块 2=类模块 3=窗体
'------------------------------------------
Private Sub ExportComponents(ByVal vbproj As Object, ByVal names As Variant, _
                             ByVal toDir As String, ByVal compType As Long)
    Dim i As Long, nm$, comp As Object, out$
    If IsEmpty(names) Then Exit Sub
    For i = LBound(names) To IBound(names)
        nm = CStr(names(i))
        On Error Resume Next
        Set comp = vbproj.VBComponents(nm)
        On Error GoTo 0
        If Not comp Is Nothing Then
            If comp.Type = compType Then
                Select Case compType
                    Case 1: out = toDir & "\" & nm & ".bas"
                    Case 2: out = toDir & "\" & nm & ".cls"
                    Case 3: out = toDir & "\" & nm & ".frm" ' .frx 会自动一起导出
                End Select
                comp.Export out
            Else
                Debug.Print "[跳过] 类型不匹配："; nm
            End If
        Else
            Debug.Print "[未找到] 组件："; nm
        End If
        Set comp = Nothing
    Next
End Sub

'------------------------------------------
' 导入：把文件夹里的 .bas/.cls/.frm 导入到目标 VBProject
'------------------------------------------
Private Sub ImportComponents(ByVal vbproj As Object, ByVal fromDir As String, ByVal pattern As String)
    Dim f As String
    f = Dir(fromDir & "\" & pattern)
    Do While Len(f) > 0
        vbproj.VBComponents.Import fromDir & "\" & f
        f = Dir
    Loop
End Sub

'------------------------------------------
' 写入“引导模块”（按钮/自安装/卸载/AutoExec）
' 说明：
' （一）加载模板时，自动创建一个“格式工具箱”工具栏（在“加载项”选项卡中出现）
' （二）按钮示例绑定到：表格_全局格式化_不依赖Selection（按需改）
' （三）提供“安装/卸载到 Word 启动文件夹”的一键入口
'------------------------------------------
'------------------------------------------
Private Sub InjectBootstrapModule(ByVal vbproj As Object)
    Dim c As Object: Set c = vbproj.VBComponents.Add(1) ' 1=标准模块
    c.name = "modBootstrap"
    
    Dim s As String
    
    '（一）启动/退出钩子
    s = s & "'（一）启动时创建按钮" & vbCrLf
    s = s & "Public Sub AutoExec()" & vbCrLf
    s = s & "    CreateToolBar" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "'（二）退出时清理按钮（可选）" & vbCrLf
    s = s & "Public Sub AutoExit()" & vbCrLf
    s = s & "    DeleteToolBar" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    '（三）创建/删除工具栏与按钮
    s = s & "'（三）创建工具栏与按钮（按钮绑定你的主功能过程）" & vbCrLf
    s = s & "Public Sub CreateToolBar()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.CommandBars(""格式工具箱"").Delete" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    Dim cb As CommandBar" & vbCrLf
    s = s & "    Set cb = Application.CommandBars.Add(Name:=""格式工具箱"", Position:=msoBarTop, Temporary:=True)" & vbCrLf
    s = s & "    Dim btn As CommandBarButton" & vbCrLf
    s = s & "    Set btn = cb.Controls.Add(Type:=msoControlButton)" & vbCrLf
    s = s & "    btn.Caption = ""表格一键格式化""" & vbCrLf
    s = s & "    btn.Style = msoButtonIconAndCaption" & vbCrLf
    s = s & "    btn.FaceId = 1085" & vbCrLf
    s = s & "    btn.OnAction = ""表格_全局格式化_不依赖Selection""" & vbCrLf
    s = s & "    cb.Visible = True" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Public Sub DeleteToolBar()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.CommandBars(""格式工具箱"").Delete" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    '（四）一键安装到启动文件夹
    s = s & "'（四）一键安装：复制本模板到 Word 启动目录（重启后自动加载）" & vbCrLf
    s = s & "Public Sub 安装到启动文件夹()" & vbCrLf
    s = s & "    Dim startup As String" & vbCrLf
    s = s & "    startup = Options.DefaultFilePath(wdStartupPath)" & vbCrLf
    s = s & "    If Len(Dir$(startup, vbDirectory)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""未找到启动文件夹："" & startup, vbExclamation: Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim src As String, dst As String, nm As String" & vbCrLf
    s = s & "    src = ThisDocument.FullName" & vbCrLf
    s = s & "    nm = CreateObject(""Scripting.FileSystemObject"").GetFileName(src)" & vbCrLf
    s = s & "    dst = startup & IIf(Right$(startup,1)=""\"" ,"""" ,""\"" ) & nm" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    FileCopy src, dst" & vbCrLf
    s = s & "    MsgBox ""已安装到："" & vbCrLf & dst & vbCrLf & ""重启 Word 生效。"", vbInformation, ""安装成功""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    MsgBox ""安装失败（"" & Err.Number & ""）："" & Err.Description, vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    '（五）一键卸载
    s = s & "'（五）一键卸载：从启动文件夹删除同名模板" & vbCrLf
    s = s & "Public Sub 卸载从启动文件夹()" & vbCrLf
    s = s & "    Dim startup As String" & vbCrLf
    s = s & "    startup = Options.DefaultFilePath(wdStartupPath)" & vbCrLf
    s = s & "    Dim nm As String: nm = CreateObject(""Scripting.FileSystemObject"").GetFileName(ThisDocument.FullName)" & vbCrLf
    s = s & "    Dim p As String: p = startup & IIf(Right$(startup,1)=""\"" ,"""" ,""\"" ) & nm" & vbCrLf
    s = s & "    If Len(Dir$(p)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""启动目录中未找到："" & vbCrLf & p, vbExclamation: Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo EH2" & vbCrLf
    s = s & "    Kill p" & vbCrLf
    s = s & "    MsgBox ""已卸载："" & vbCrLf & p & vbCrLf & ""重启 Word 生效。"", vbInformation" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH2:" & vbCrLf
    s = s & "    MsgBox ""卸载失败（"" & Err.Number & ""）："" & Err.Description, vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf
    
    c.codeModule.AddFromString s
End Sub

