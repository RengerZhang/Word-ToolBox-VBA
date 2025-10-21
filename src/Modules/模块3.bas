Attribute VB_Name = "模块3"
Option Explicit

'==========================================================
' 诊断：Normal.dotm 备份权限
' 作用：分别检测 源文件是否可读、目标文件夹是否可写，并做一次模拟复制测试
' 使用：将本宏粘贴到模块，运行 诊断_Normal备份权限()
' 目标路径已写死：C:\Users\Tony Zhang\Desktop\VBA源代码备份\
'==========================================================
Sub 诊断_Normal备份权限()
    On Error GoTo ERR_HANDLER
    
    '（一）定位源与目标
    Dim src As String, srcDir As String
    Dim targetDir As String
    src = Application.NormalTemplate.FullName                     ' 当前使用的 Normal.dotm
    srcDir = Left$(src, InStrRev(src, "\") - 1)                   ' 源文件所在目录
    targetDir = "C:\Users\Tony Zhang\Desktop\VBA源代码备份\"      ' 目标备份目录（写死）
    
    '（二）保存 Normal.dotm，避免“未落盘”导致读失败
    Application.NormalTemplate.Save
    
    '（三）逐项检测
    Dim srcExists As Boolean, srcReadable As Boolean
    Dim tgtExists As Boolean, tgtWritable As Boolean
    Dim copyOk As Boolean, copyMsg As String
    Dim testDst As String, stamp As String
    
    srcExists = FileExists(src)
    If srcExists Then srcReadable = CanReadFile(src)
    
    tgtExists = FolderExists(targetDir)
    If tgtExists Then tgtWritable = HasWriteAccess(targetDir)
    
    '（四）尝试“模拟复制”到目标目录（不影响你的正式备份，复制后立刻删除）
    If srcReadable And tgtWritable Then
        stamp = Format(Now, "yyyymmdd_HHNNSS")
        testDst = AddSlash(targetDir) & "Normal_permtest_" & stamp & ".dotm"
        On Error Resume Next
        FileCopy src, testDst
        If Err.Number = 0 Then
            copyOk = True
            ' 复制成功后尝试删除测试文件
            Kill testDst
        Else
            copyOk = False
            copyMsg = "模拟复制失败（" & Err.Number & "）： " & Err.Description
        End If
        On Error GoTo ERR_HANDLER
    End If
    
    '（五）汇总报告
    Dim markOK As String, markNG As String
    markOK = "?"
    markNG = "?"
    
    Dim REPORT As String
    REPORT = ""
    REPORT = REPORT & "【源文件】" & vbCrLf
    REPORT = REPORT & "路径： " & src & vbCrLf
    REPORT = REPORT & "所在文件夹： " & srcDir & vbCrLf
    REPORT = REPORT & "存在： " & IIf(srcExists, markOK, markNG) & vbCrLf
    REPORT = REPORT & "可读： " & IIf(srcReadable, markOK, markNG) & vbCrLf & vbCrLf
    
    REPORT = REPORT & "【目标文件夹】" & vbCrLf
    REPORT = REPORT & "路径： " & targetDir & vbCrLf
    REPORT = REPORT & "存在： " & IIf(tgtExists, markOK, markNG) & vbCrLf
    REPORT = REPORT & "可写： " & IIf(tgtWritable, markOK, markNG) & vbCrLf & vbCrLf
    
    REPORT = REPORT & "【模拟复制】" & vbCrLf
    If srcReadable And tgtWritable Then
        REPORT = REPORT & "结果： " & IIf(copyOk, "成功 " & markOK, "失败 " & markNG) & vbCrLf
        If Not copyOk Then REPORT = REPORT & copyMsg & vbCrLf
    Else
        REPORT = REPORT & "未执行（因源不可读或目标不可写）" & vbCrLf
    End If
    
    '（六）给出判断结论（简化判断逻辑）
    REPORT = REPORT & vbCrLf & "【结论】" & vbCrLf
    If Not srcReadable And Not tgtWritable Then
        REPORT = REPORT & "源文件不可读 + 目标目录不可写。两边均存在权限/访问问题。" & vbCrLf
    ElseIf Not srcReadable Then
        REPORT = REPORT & "源文件（Normal.dotm 或其所在文件夹）无读取权限或被占用。" & vbCrLf
    ElseIf Not tgtWritable Then
        REPORT = REPORT & "目标文件夹无写入权限（或被系统安全策略拦截）。" & vbCrLf
    ElseIf srcReadable And tgtWritable And Not copyOk Then
        REPORT = REPORT & "读/写检测通过，但复制仍失败，多半是安全策略（如“受控文件夹访问”）或安全软件拦截。" & vbCrLf
    Else
        REPORT = REPORT & "源可读、目标可写、模拟复制正常。此前失败可能为临时占用或路径问题。" & vbCrLf
    End If
    
    MsgBox REPORT, IIf(copyOk, vbInformation, vbExclamation), "Normal.dotm 备份权限诊断"
    Exit Sub

ERR_HANDLER:
    MsgBox "诊断过程出错（" & Err.Number & "）： " & Err.Description, vbCritical, "错误"
End Sub

'------------------------------------------
'（工具1）文件是否存在
'------------------------------------------
Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir$(filePath, vbNormal Or vbReadOnly Or vbSystem Or vbHidden)) > 0)
End Function

'------------------------------------------
'（工具2）文件夹是否存在
'------------------------------------------
Private Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then FolderExists = False: Exit Function
    FolderExists = (Len(Dir$(AddSlash(folderPath), vbDirectory)) > 0)
End Function

'------------------------------------------
'（工具3）测试文件可读：尝试打开为只读
'------------------------------------------
Private Function CanReadFile(ByVal filePath As String) As Boolean
    On Error GoTo NG
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary Access Read As #ff
    Close #ff
    CanReadFile = True
    Exit Function
NG:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    CanReadFile = False
End Function

'------------------------------------------
'（工具4）测试目录可写：尝试创建并删除临时文件
'------------------------------------------
Private Function HasWriteAccess(ByVal folderPath As String) As Boolean
    On Error GoTo NG
    Dim testFile As String, ff As Integer
    folderPath = AddSlash(folderPath)
    testFile = folderPath & "~write_test_" & Format(Now, "yyyymmdd_HHNNSS") & ".tmp"
    ff = FreeFile
    Open testFile For Output As #ff
    Print #ff, "test"
    Close #ff
    Kill testFile
    HasWriteAccess = True
    Exit Function
NG:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    HasWriteAccess = False
End Function

'------------------------------------------
'（工具5）统一补全目录反斜杠
'------------------------------------------
Private Function AddSlash(ByVal p As String) As String
    If Len(p) = 0 Then
        AddSlash = ""
    ElseIf Right$(p, 1) = "\" Then
        AddSlash = p
    Else
        AddSlash = p & "\"
    End If
End Function

