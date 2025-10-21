Attribute VB_Name = "模块2"
Option Explicit

'==========================================================
' 独立测试：清除“表格标题”段的手工编号前缀（两步）
' 规则：
'  - Step A：若段落以“表”开头，删除从“表”到“第一个中文字符”为止
'  - Step B：若 A 未命中，再用“宽容但不失控”的正则清除常见编号形态：
'            表 [可选连字符] 数字 [点分数字]* [可选 - 数字]
'            兼容全/半角点（. ． 。）与多种连字符（- － – —）
' 仅遍历正文（wdMainTextStory）
' 默认仅处理样式【表格标题】，若样式不存在则退化为处理“所有以‘表’开头的段落”
'==========================================================
Sub 测试_清除表题手工编号_独立()
    Const 仅处理指定样式 As Boolean = True
    Const 表题样式名 As String = "表格标题"
    
    Dim doc As Document: Set doc = ActiveDocument
    Dim capStyle As Style, useStyleFilter As Boolean
    Dim p As Paragraph, r As Range
    Dim oldTxt As String, newTxt As String
    Dim total As Long, touched As Long, skipped As Long, examples As Long
    
    '——样式过滤：若不存在目标样式，则退化为“不过滤样式”
    useStyleFilter = 仅处理指定样式
    If useStyleFilter Then
        On Error Resume Next
        Set capStyle = doc.Styles(表题样式名)
        On Error GoTo 0
        If capStyle Is Nothing Then useStyleFilter = False
    End If
    
    '——可选：把所有修改放入一个撤销块（新版本 Word 支持）
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "清除表题手工编号"
    On Error GoTo 0
    
    For Each p In doc.Paragraphs
        ' 仅正文
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextPara
        
        ' 样式过滤（如果启用）
        If useStyleFilter Then
            On Error Resume Next
            If p.Range.Style.nameLocal <> 表题样式名 Then GoTo NextPara
            On Error GoTo 0
        End If
        
        ' 只处理“以‘表’开头”的段落（含前导空格/全角空格）
        oldTxt = 清理段首可见文本(p.Range.text)
        If Len(oldTxt) = 0 Or Left$(oldTxt, 1) <> "表" Then GoTo NextPara
        
        '——目标子范围（不含段尾标记）
        Set r = p.Range.Duplicate
        If r.Characters.Count > 1 Then r.MoveEnd wdCharacter, -1
        
        '——Step A：表… → 删到第一个中文字符
        newTxt = 去除表题旧前缀_到第一个中文(r.text)
        
        '——Step B：若 A 未改变文本，再用正则兜底“表[可选连字符]数字[点分]*[可选-数字]”
        '   解释本正则：
        '   ^\s*                 —— 从段首开始，允许若干空白（含全角空格已在上游转为半角）
        '   表\s*                —— “表”后允许若干空格
        '   [-－–—]?             —— 可选连字符（覆盖常见的 -、全角－、短横–、长横—）
        '   \s*\d+               —— 若干空格后至少一位数字（严格防止“表A-1”被误匹配）
        '   (?:\s*[\.．。]\s*\d+)* —— 0~多次“点 + 数字”（点可为半角.、全角．、中文。）
        '   \s*                  —— 可选空格
        '   (?:[-－–—]\s*\d+)?   —— 可选“连字符 + 数字”（顺序号，如 -1）
        '   \s*                  —— 吃掉编号后面的空格
        If newTxt = r.text Then
            newTxt = 正则替换( _
                newTxt, _
                "^\s*表\s*[-－–—]?\s*\d+(?:\s*[\.．。]\s*\d+)*\s*(?:[-－–—]\s*\d+)?\s*", _
                "" _
            )
            newTxt = LTrim$(newTxt)
        End If
        
        total = total + 1
        If newTxt <> r.text Then
            r.text = newTxt
            touched = touched + 1
            ' 打印前 8 条修改示例到“立即窗口”
            If examples < 8 Then
                Debug.Print "★改前："; oldTxt
                Debug.Print " 改后："; 清理段首可见文本(newTxt)
                examples = examples + 1
            End If
        Else
            skipped = skipped + 1
        End If
        
NextPara:
    Next p
    
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    
    MsgBox "处理完成：" & vbCrLf & _
           "候选段落（以“表”开头）：" & total & vbCrLf & _
           "已清除前缀：" & touched & vbCrLf & _
           "未变更（无匹配）：" & skipped & vbCrLf & vbCrLf & _
           "提示：按 Ctrl+G 打开“立即窗口”可查看示例。", vbInformation
End Sub

'——Step A：若以“表”开头，从“表”删到“第一个中文字符”
Private Function 去除表题旧前缀_到第一个中文(ByVal s As String) As String
    Dim i As Long, ch As String, hit As Boolean
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = LTrim$(s)
    
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
        去除表题旧前缀_到第一个中文 = s   ' 没找到中文，交给正则兜底
    End If
End Function

'——判断是否中文（修正 AscW 负数问题；CJK 基本区 + 扩展A）
Private Function 是否中文字符(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) = 0 Then 是否中文字符 = False: Exit Function
    code = AscW(ch)
    If code < 0 Then code = code + &H10000  ' ★关键：把有符号值归一到 0..65535
    是否中文字符 = ((code >= &H4E00 And code <= &H9FFF) Or (code >= &H3400 And code <= &H4DBF))
End Function

'——正则：单次替换（大小写不敏感）
Private Function 正则替换(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim r As Object: Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False          ' 只替换段首的那一处前缀；其余由“多次运行”策略处理
    r.pattern = pat
    正则替换 = r.Replace(s, rep)
End Function

'——清洗：去段尾、去单元格结束符、全角空格→半角并 Trim
Private Function 清理段首可见文本(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    清理段首可见文本 = Trim$(s)
End Function


