Attribute VB_Name = "编号_6_带进度的去除编号"
Option Explicit

'==========================================================
' ④ 删除手工编号（带进度条 + 循环确认）
' 说明：
'   - 目标样式 & 删除规则：来自“配置中心”
'   - 仅删除段首手工编号，不影响多级自动编号
'   - 遍历时跳过：非正文Story、表格内段落、目录(TOC)区域
'   - 每一轮结束，如果“已清除>0”，询问是否继续下一轮
'==========================================================
Public Sub 去除手工编号_使用进度窗体()
    Dim doc As Document: Set doc = ActiveDocument
    Dim scope As Range              ' ← 本次处理范围：选区 或 全文
    Dim scopeInfo As String
    Dim backupPath As String
    Dim targetStyles As Variant
    Dim patterns As Variant
    Dim tocZones As Collection
    Dim cand As Long
    Dim passNo As Long
    Dim touched As Long, skipped As Long, total As Long, allTouched As Long
    Dim ans As VbMsgBoxResult

    '——1) 处理范围：若“选中且在正文Story”，仅处理选中段落；否则全文
    If Selection.Type <> wdSelectionIP And Selection.Range.StoryType = wdMainTextStory Then
        Set scope = Selection.Range.Duplicate
        scopeInfo = "范围：选中段落"
    Else
        Set scope = doc.content.Duplicate
        scopeInfo = "范围：全文"
    End If

    '——2) 运行前备份（同目录）
    backupPath = 备份当前文档(doc)
    If Len(backupPath) > 0 Then Debug.Print "已备份到: " & backupPath

    '——3) 从“配置中心”读取：样式集合 & 删除规则
    targetStyles = 获取样式名数组(True)          ' 只返回文档中已存在的目标样式
    patterns = 生成删除编号规则集()               ' 动态生成删除模式

    '——4) 收集 TOC 区域（用于跳过目录内容）
    Set tocZones = 构建TOC区域集(doc)

    '——5) 统计候选段（在 scope 内）
    cand = 统计候选段数_删除编号(scope, targetStyles, tocZones)

    '——6) 打开进度窗体
    With progressForm
        .caption = "删除手工编号"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = scopeInfo & "；候选段落：" & cand & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "删除手工编号（循环）"
    On Error GoTo 0

    Do
        passNo = passNo + 1
        progressForm.UpdateProgressBar 0, "—— 第 " & passNo & " 轮开始 ——"

        ' 每轮开头重新统计（上一轮可能改动了样式/文本）
        cand = 统计候选段数_删除编号(scope, targetStyles, tocZones)
        If cand = 0 Then
            progressForm.UpdateProgressBar 200, "没有候选段落，直接结束。"
            Exit Do
        End If

        total = 0: touched = 0: skipped = 0
        执行一轮删除 scope, targetStyles, patterns, tocZones, cand, total, touched, skipped

        progressForm.UpdateProgressBar 200, _
            "第 " & passNo & " 轮小结：候选=" & total & "，已清除=" & touched & "，未变更=" & skipped

        allTouched = allTouched + touched
        If progressForm.stopFlag Then
            MsgBox "已手动终止。累计清除：" & allTouched & " 处。", vbExclamation
            Exit Do
        End If

        If touched = 0 Then
            MsgBox "删除手工编号完成（本轮无可清除项）。" & vbCrLf & _
                   "累计清除：" & allTouched & " 处。", vbInformation
            Exit Do
        Else
            ans = MsgBox("本轮已清除 " & touched & " 处编号前缀。" & vbCrLf & _
                         "可能仍有残留，是否继续下一轮？", _
                         vbYesNo + vbQuestion, "继续清除？")
            If ans = vbNo Then Exit Do
            progressForm.FrameProgress.width = 0
            progressForm.LabelPercentage.caption = "0%"
        End If

        If passNo >= 5 Then
            ans = MsgBox("已运行 5 轮，是否仍继续？", vbYesNo + vbExclamation)
            If ans = vbNo Then Exit Do
        End If
    Loop

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True

    If Not progressForm.stopFlag Then
        progressForm.UpdateProgressBar 200, "完成。累计清除：" & allTouched
        MsgBox "删除手工编号：处理结束。累计清除 " & allTouched & " 处。", vbInformation
    End If
End Sub



' ----------------------------------------------------------
' 单轮执行：对所有目标样式做删除尝试（仅段首）
'  - 跳过：非正文、表格内、TOC 区域
'  - 每命中一段：依次跑所有删除规则（单次替换），再清段首空格
'  - 为防止“末段死循环”，每次显式把 rng 跳到下一段
' candTotal 用于进度条；输出 total/touched/skipped
' ----------------------------------------------------------
' 单轮执行：逐段遍历，不再用 Find，避免末段越界
Private Sub 执行一轮删除(ByVal scope As Range, _
                       ByVal targetStyles As Variant, _
                       ByVal patterns As Variant, _
                       ByVal tocZones As Collection, _
                       ByVal candTotal As Long, _
                       ByRef total As Long, _
                       ByRef touched As Long, _
                       ByRef skipped As Long)

    Dim p As Paragraph
    Dim contentRng As Range
    Dim originalText As String, newText As String
    Dim pat As Variant, sty As String
    Dim processed As Long, examples As Long

    For Each p In scope.Paragraphs
        ' 过滤：正文 / 非表格 / 非目录
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextP
        If 在表格内(p.Range) Then GoTo NextP
        If 在TOC区域内(p.Range, tocZones) Then GoTo NextP

        On Error Resume Next
        sty = p.Range.Style.nameLocal
        On Error GoTo 0
        If Not 样式在列表中(sty, targetStyles) Then GoTo NextP

        total = total + 1

        ' 取可编辑内容（不含段尾标记）
        Set contentRng = p.Range.Duplicate
        If contentRng.Characters.Count > 1 Then contentRng.MoveEnd wdCharacter, -1

        originalText = contentRng.text
        newText = originalText

        ' 依次套用所有“删除段首编号”的正则（每条只替首处）
        For Each pat In patterns
            newText = 正则替换_一次(newText, CStr(pat), "")
        Next
        ' 清段首残留空格（含全角）
        newText = 正则替换_一次(newText, "^[ 　]+", "")

        If newText <> originalText Then
            contentRng.text = newText
            touched = touched + 1
            If examples < 6 Then
                progressForm.UpdateProgressBar 当前进度像素(processed, IIf(candTotal = 0, 1, candTotal)), _
                    "★改前：" & Left$(originalText, 80) & vbCrLf & " 改后：" & Left$(newText, 80)
                examples = examples + 1
            End If
        Else
            skipped = skipped + 1
        End If

        processed = processed + 1
        progressForm.UpdateProgressBar 当前进度像素(processed, IIf(candTotal = 0, 1, candTotal)), _
            "进度：" & processed & "/" & candTotal
NextP:
        DoEvents
    Next p
End Sub


' 候选段：正文Story、非表格、非TOC、样式 ∈ 目标样式集
Private Function 统计候选段数_删除编号(ByVal scope As Range, _
                                    ByVal targetStyles As Variant, _
                                    ByVal tocZones As Collection) As Long
    Dim p As Paragraph, n As Long, sty As String

    For Each p In scope.Paragraphs
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextP
        If 在表格内(p.Range) Then GoTo NextP
        If 在TOC区域内(p.Range, tocZones) Then GoTo NextP
        On Error Resume Next
        sty = p.Range.Style.nameLocal
        On Error GoTo 0
        If 样式在列表中(sty, targetStyles) Then n = n + 1
NextP:
    Next
    统计候选段数_删除编号 = n
End Function

Private Function 样式在列表中(ByVal sty As String, ByVal arr As Variant) As Boolean
    Dim v As Variant
    For Each v In arr
        If StrComp(sty, CStr(v), vbTextCompare) = 0 Then 样式在列表中 = True: Exit Function
    Next
End Function

' 是否在表格里（双保险）
Private Function 在表格内(ByVal r As Range) As Boolean
    On Error Resume Next
    在表格内 = r.Information(wdWithInTable)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If Not 在表格内 Then 在表格内 = (r.Tables.Count > 0)
End Function

'——构建 TOC 字段结果区域集合
Private Function 构建TOC区域集(ByVal doc As Document) As Collection
    Dim zones As New Collection
    Dim f As Field, codeTxt As String
    On Error Resume Next
    For Each f In doc.Fields
        codeTxt = ""
        codeTxt = f.code.text
        If (f.Type = wdFieldTOC) Or (InStr(1, UCase$(codeTxt), "TOC", vbTextCompare) > 0) Then
            zones.Add f.Result.Duplicate
        End If
    Next f
    Set 构建TOC区域集 = zones
End Function

'——判定 Range 是否完全落在任一 TOC 结果区域内
Private Function 在TOC区域内(ByVal r As Range, ByVal zones As Collection) As Boolean
    Dim z As Range
    If zones Is Nothing Then Exit Function
    On Error Resume Next
    For Each z In zones
        If (r.Start >= z.Start) And (r.End <= z.End) Then 在TOC区域内 = True: Exit Function
    Next z
End Function

' 正则：单次替换（仅首处）
Private Function 正则替换_一次(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True: rx.Global = False: rx.pattern = pat
    正则替换_一次 = rx.Replace(s, rep)
End Function

' 正则：仅判定
Private Function 正则命中(ByVal s As String, ByVal pat As String) As Boolean
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True: rx.Global = False: rx.pattern = pat
    正则命中 = rx.TEST(s)
End Function

' 备份到同目录
Private Function 备份当前文档(ByVal doc As Document) As String
    On Error GoTo EH
    Dim baseName As String, ext As String, bak As String, folder As String, ts As String
    ts = Format(Now, "yyyymmdd_hhnnss")
    If Len(doc.name) > 0 Then
        baseName = doc.name
        If InStrRev(baseName, ".") > 0 Then
            ext = mid$(baseName, InStrRev(baseName, "."))
            baseName = Left$(doc.name, InStrRev(doc.name, ".") - 1)
        Else
            ext = ".docx"
        End If
    Else
        baseName = "未命名文档": ext = ".docx"
    End If
    folder = IIf(doc.path = "", CurDir$, doc.path)
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    bak = folder & baseName & "_备份_" & ts & ext
    doc.SaveCopyAs FileName:=bak
    备份当前文档 = bak
    Exit Function
EH:
    备份当前文档 = ""
End Function

' 进度像素（你的窗体满条 200px）
Private Function 当前进度像素(ByVal done As Long, ByVal total As Long) As Long
    If total <= 0 Then
        当前进度像素 = 0
    Else
        当前进度像素 = CLng(200# * done / total)
        If 当前进度像素 < 0 Then 当前进度像素 = 0
        If 当前进度像素 > 200 Then 当前进度像素 = 200
    End If
End Function

