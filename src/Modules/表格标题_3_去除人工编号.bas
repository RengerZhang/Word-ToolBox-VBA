Attribute VB_Name = "表格标题_3_去除人工编号"
Option Explicit
Private Const 调试_打印码点 As Boolean = False


' ==========================================================
' 表标题手工编号清理（循环 + 进度条窗体）
' 规则：
'   （一）仅处理“正文”Story中、段首（经预清理后）以“表”开头的段落
'   （二）Step A：从“表”删到“第一个中文字符”为止（可吃掉题注产生的控制符）
'   （三）Step B（兜底）：正则清“表[可选连字符]数字[点分]*[可选-数字]”
'   （四）每轮结束：若已清除>0，弹窗询问是否继续；=0 则提示“清除完毕”
'   （五）进度与日志：使用你已有的 ProgressForm（无须改窗体）
' ==========================================================

Public Sub 清除表题手工编号_使用进度窗体1()
    Const 仅处理指定样式 As Boolean = True
    Const 表题样式名 As String = "表格标题"

    Dim passNo As Long
    Dim total As Long, touched As Long, skipped As Long
    Dim allTouched As Long
    Dim ans As VbMsgBoxResult
    Dim doc As Document: Set doc = ActiveDocument
    Dim useStyleFilter As Boolean
    Dim capStyle As Style

    ' （一）样式过滤判定
    useStyleFilter = 仅处理指定样式
    If useStyleFilter Then
        On Error Resume Next
        Set capStyle = doc.Styles(表题样式名)
        On Error GoTo 0
        If capStyle Is Nothing Then useStyleFilter = False
    End If

    ' （二）显示进度窗体（无模式）
    With progressForm
        .caption = "表标题前缀清除"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = "准备中……" & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With

'    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "表标题前缀清除（循环）"
    On Error GoTo 0

    Do
        passNo = passNo + 1
        progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & _
            "—— 第 " & passNo & " 轮开始 ——" & vbCrLf

        ' （三）统计候选数量（用于计算进度）
        Dim cand As Long
        cand = 统计候选段数(useStyleFilter, 表题样式名)

        If cand = 0 Then
            progressForm.UpdateProgressBar 200, "本轮没有以“表”开头的候选段落。"
            MsgBox "表标题前缀清除完毕（未发现候选）。" & vbCrLf & _
                   "累计已清除：" & allTouched & " 处。", vbInformation
            Exit Do
        End If

        ' （四）执行单轮清理（带进度与日志）
        执行一轮清除 cand, useStyleFilter, 表题样式名, total, touched, skipped
        allTouched = allTouched + touched

        ' （五）轮次小结
        Dim summary As String
        summary = "第 " & passNo & " 轮完成：候选=" & total & "，已清除=" & touched & "，未变更=" & skipped
        progressForm.UpdateProgressBar 200, summary

        ' （六）结束条件 / 继续提示
        If progressForm.stopFlag Then
            MsgBox "已手动终止。累计已清除：" & allTouched & " 处。", vbExclamation
            Exit Do
        End If

        If touched = 0 Then
            MsgBox "表标题前缀清除完毕（本轮无可清除项）。" & vbCrLf & _
                   "累计已清除：" & allTouched & " 处。", vbInformation
            Exit Do
        Else
            ans = MsgBox("本轮已清除 " & touched & " 处前缀。" & vbCrLf & _
                         "前缀可能尚未完全清除，是否继续下一轮？", _
                         vbYesNo + vbQuestion, "继续清除？")
            If ans = vbNo Then Exit Do
            ' 重置进度条
            progressForm.FrameProgress.width = 0
            progressForm.LabelPercentage.caption = "0%"
        End If

        ' 安全阀：避免极端循环
        If passNo >= 5 Then
            ans = MsgBox("已运行 5 轮，是否仍要继续？", vbYesNo + vbExclamation)
            If ans = vbNo Then Exit Do
        End If
    Loop

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True

'    ' 收尾：隐藏窗体
'    On Error Resume Next
'    Unload progressForm
'    On Error GoTo 0
End Sub


' ----------------------------------------------------------
' 执行单轮清理（带进度条）
' candTotal: 统计到的候选总数（用于进度计算）
' useStyleFilter / targetStyleName: 是否仅处理【表格标题】样式
' 输出：total/touched/skipped（本轮统计）
' ----------------------------------------------------------
Private Sub 执行一轮清除(ByVal candTotal As Long, _
                        ByVal useStyleFilter As Boolean, _
                        ByVal targetStyleName As String, _
                        ByRef total As Long, _
                        ByRef touched As Long, _
                        ByRef skipped As Long)

    Dim doc As Document: Set doc = ActiveDocument
    Dim p As Paragraph, r As Range
    Dim oldTxt As String, newTxt As String
    Dim processed As Long, progressPx As Long
    Dim examples As Long

    total = 0: touched = 0: skipped = 0

    For Each p In doc.Paragraphs
        If progressForm.stopFlag Then Exit For
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextPara

'        ' ——（1）样式过滤（如启用）
'        If useStyleFilter Then
'            On Error Resume Next
'            If p.Range.Style.NameLocal <> targetStyleName Then GoTo NextPara
'            On Error GoTo 0
'        End If

        ' ——（2）候选判定：段首（强预清理后）需以“表”开头
        oldTxt = 强预清理_段首(p.Range.text)
        If Len(oldTxt) = 0 Or Left$(oldTxt, 1) <> "表" Then GoTo NextPara
        
        ' （新增）码点调试：原文 & 净化后
        If 调试_打印码点 Then
            Debug.Print String(48, "-")
            Debug.Print "候选#"; processed + 1; "/", candTotal
            Debug.Print "原文："; Left$(p.Range.text, 120)
            Debug.Print "原文码点："; 列出前N码点(p.Range.text, 60)
            Debug.Print "净化："; oldTxt
            Debug.Print "净化码点："; 列出前N码点(oldTxt, 60)
        End If

        total = total + 1
        processed = processed + 1

        ' ——（3）目标子范围（不含段尾标记）
        Set r = p.Range.Duplicate
        If r.Characters.Count > 1 Then r.MoveEnd wdCharacter, -1

        ' ——（4）Step A：从“表”删到“首个中文”
        newTxt = 去除表题旧前缀_到第一个中文(r.text)

        ' ——（5）Step B：若 A 未改变文本，用正则兜底
        If newTxt = r.text Then
            newTxt = 正则替换(newTxt, _
                    "^\s*表\s*[-－–—]?\s*\d+(?:\s*[\.．。]\s*\d+)*\s*(?:[-－–—]\s*\d+)?(?:\s+\d+)?\s*", _
                    "")
            ' 兜底后，再清掉段首噪音直到首个中文（解决题注控制符残留）
            newTxt = 去噪直至首个中文(newTxt)
        End If

        ' ——（6）回写 & 打样例日志
        If newTxt <> r.text Then
            r.text = newTxt
            touched = touched + 1
            If examples < 6 Then
                progressForm.UpdateProgressBar 当前进度像素(processed, candTotal), _
                    "★改前：" & Left$(oldTxt, 80) & vbCrLf & " 改后：" & Left$(强预清理_段首(newTxt), 80)
                examples = examples + 1
            Else
                progressForm.UpdateProgressBar 当前进度像素(processed, candTotal), _
                    "已处理：" & processed & "/" & candTotal
            End If
        Else
            skipped = skipped + 1
            progressForm.UpdateProgressBar 当前进度像素(processed, candTotal), _
                "未变更：" & Left$(oldTxt, 80)
        End If

NextPara:
        DoEvents
    Next p
End Sub


Private Function 统计候选段数(ByVal useStyleFilter As Boolean, ByVal targetStyleName As String) As Long
    Dim doc As Document: Set doc = ActiveDocument
    Dim sty As Style
    Dim scope As Range, rng As Range
    Dim cnt As Long, nextPos As Long

    ' 1) 仅处理【表格标题】样式；若不存在，提示并退出
    On Error Resume Next
    Set sty = doc.Styles("表格标题")
    On Error GoTo 0
    If sty Is Nothing Then
        MsgBox "请先执行表格标题格式匹配！", vbExclamation
        统计候选段数 = 0
        Exit Function
    End If

    ' 2) 处理范围：若当前有选区且在正文，只统计选区；否则统计全文正文
    If Selection.Type <> wdSelectionIP And Selection.Range.StoryType = wdMainTextStory Then
        Set scope = Selection.Range.Duplicate
    Else
        Set scope = doc.StoryRanges(wdMainTextStory).Duplicate
    End If

    ' 3) 用 Find 按样式统计；每次命中后跳到该段末，避免重复/卡滞
    Set rng = scope.Duplicate
    With rng.Find
        .ClearFormatting
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = sty

        Do While .Execute
            cnt = cnt + 1
            nextPos = rng.Paragraphs(1).Range.End   ' 跳到命中段落的段尾
            If nextPos >= scope.End Then Exit Do    ' 到达范围末尾则结束
            rng.SetRange Start:=nextPos, End:=scope.End
        Loop
    End With

    统计候选段数 = cnt
End Function



' ----------------------------------------------------------
' 进度像素：你的窗体200px为满条 → 像素 = 200 * done / total
' ----------------------------------------------------------
Private Function 当前进度像素(ByVal done As Long, ByVal total As Long) As Long
    If total <= 0 Then
        当前进度像素 = 0
    Else
        当前进度像素 = CLng(200# * done / total)
        If 当前进度像素 < 0 Then 当前进度像素 = 0
        If 当前进度像素 > 200 Then 当前进度像素 = 200
    End If
End Function


' ================= 核心规则函数 =================

'（一）Step A：若段首（强预清理后）以“表”开头，把“表”→“首个中文”之间全部删除
Private Function 去除表题旧前缀_到第一个中文(ByVal s As String) As String
    Dim i As Long, ch As String, hit As Boolean
    s = 强预清理_段首(s)
    If Len(s) = 0 Or Left$(s, 1) <> "表" Then
        去除表题旧前缀_到第一个中文 = s
        Exit Function
    End If
    For i = 2 To Len(s)
        ch = mid$(s, i, 1)
        If 是否中文字符(ch) Then hit = True: Exit For
    Next i
    If hit Then
        去除表题旧前缀_到第一个中文 = LTrim$(mid$(s, i))
    Else
        去除表题旧前缀_到第一个中文 = s   ' 没找到中文，交给 Step B + 去噪兜底
    End If
End Function

'（二）Step B 之后的去噪：把段首所有“空白/控制符/连字符/点号/数字”剥离到首个中文
Private Function 去噪直至首个中文(ByVal s As String) As String
    Dim i As Long, ch As String
    s = 强预清理_段首(s)
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If 是否中文字符(ch) Then 去噪直至首个中文 = LTrim$(mid$(s, i)): Exit Function
    Next i
    去噪直至首个中文 = s
End Function


' ================= 清洗/判断工具 =================

'（工具）段首强预清理：去最常见“不可见噪音”，再左 Trim
Private Function 强预清理_段首(ByVal s As String) As String
    Dim i As Long, out As String, ch As String, cp As Long
    ' 快速替换：段尾、单元格结束、全角空格、NBSP、Tab
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = Replace$(s, ChrW(&HA0), " ")
    s = Replace$(s, vbTab, " ")
    ' 零宽/方向控制
    s = Replace$(s, ChrW(&H200B), "")
    s = Replace$(s, ChrW(&H200C), "")
    s = Replace$(s, ChrW(&H200D), "")
    s = Replace$(s, ChrW(&HFEFF), "")
    s = Replace$(s, ChrW(&H200E), "")
    s = Replace$(s, ChrW(&H200F), "")
    s = Replace$(s, ChrW(&H202A), "")
    s = Replace$(s, ChrW(&H202B), "")
    s = Replace$(s, ChrW(&H202C), "")
    s = Replace$(s, ChrW(&H202D), "")
    s = Replace$(s, ChrW(&H202E), "")
    ' 记录分隔符等控制符（含 U+001E）
    s = Replace$(s, ChrW(&H1E), "")
    ' 保险：滤除 <32 的控制字符
    out = ""
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        cp = AscW(ch)
        If cp < 0 Then cp = cp + &H10000
        If cp >= 32 Then out = out & ch
    Next i
    强预清理_段首 = LTrim$(out)
End Function

'（工具）判断是否中文（修正 AscW 负数；CJK 基本区 + 扩展A）
Private Function 是否中文字符(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) = 0 Then 是否中文字符 = False: Exit Function
    code = AscW(ch)
    If code < 0 Then code = code + &H10000
    是否中文字符 = ((code >= &H4E00 And code <= &H9FFF) Or (code >= &H3400 And code <= &H4DBF))
End Function

'（工具）正则：单次替换（只替段首一处；其余靠“循环多轮”）
Private Function 正则替换(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Global = False
    rx.pattern = pat
    正则替换 = rx.Replace(s, rep)
End Function

'（工具）列出前 N 个字符的 Unicode 码点（已修正 AscW 负数），形如 "U+8868 U+002D ..."
Private Function 列出前N码点(ByVal s As String, ByVal n As Long) As String
    Dim i As Long, out As String, cp As Long, ch As String
    Dim m As Long: m = IIf(Len(s) < n, Len(s), n)
    For i = 1 To m
        ch = mid$(s, i, 1)
        cp = AscW(ch)
        If cp < 0 Then cp = cp + &H10000
        out = out & "U+" & Right$("0000" & Hex$(cp), 4) & " "
    Next
    列出前N码点 = Trim$(out)
End Function

