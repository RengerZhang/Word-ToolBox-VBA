Attribute VB_Name = "MOD_公共工具函数"
Option Explicit

' ---------- 工具函数：中文字号文本转磅值（公共，仅此一个） ----------
' 支持中文字号（如 五号、小四 等），也允许直接输入数字磅值（如 12、10.5）
Public Function GetFontSizePt(sizeText As String) As Single
    Dim sizeMap As Object
    Set sizeMap = CreateObject("Scripting.Dictionary")
    With sizeMap
        .Add "初号", 42
        .Add "小初", 36
        .Add "一号", 26
        .Add "小一", 24
        .Add "二号", 22
        .Add "小二", 18
        .Add "三号", 16
        .Add "小三", 15
        .Add "四号", 14
        .Add "小四", 12
        .Add "五号", 10.5
        .Add "小五", 9
        .Add "六号", 7.5
        .Add "小六", 6.5
    End With

    If sizeMap.exists(sizeText) Then
        GetFontSizePt = sizeMap(sizeText)
    ElseIf IsNumeric(sizeText) Then
        GetFontSizePt = CSng(sizeText)
    Else
        GetFontSizePt = -1
    End If
End Function


'==========================================================
' 配置中心：从“自动编号（步骤三）”里读取级别参数，
' 自动产出：
'   1) 目标样式名数组（供②/④使用）
'   2) 删除手工编号的规则集（正则数组，供④使用）
'   3) 标题匹配的规则集（pattern→style 映射，供②使用）
'
' 依赖：
'   - 步骤三里的 Public Function 获取所有级别参数() As Variant
'     返回二维数组：(级别, 列)，列 = 1:样式名, 2:编号格式, 3:编号样式, 4:对齐位置
'==========================================================

'——— 读取“样式名”数组（默认只返回文档中已存在的样式）
Public Function 获取样式名数组(Optional onlyExisting As Boolean = True) As Variant
    Dim cfg As Variant, i As Long, N As Long
    Dim buf() As String
    Dim sty As Style, name As String
    
    cfg = 获取所有级别参数()  ' 来自【步骤三】的权威参数表
    ReDim buf(1 To UBound(cfg, 1))
    N = 0
    
    For i = 1 To UBound(cfg, 1)
        name = CStr(cfg(i, 1))
        If onlyExisting Then
            On Error Resume Next
            Set sty = ActiveDocument.Styles(name)
            On Error GoTo 0
            If Not sty Is Nothing Then
                N = N + 1: buf(N) = name
                Set sty = Nothing
            End If
        Else
            N = N + 1: buf(N) = name
        End If
    Next i
    
    If N = 0 Then
        获取样式名数组 = Array() ' 空
    Else
        ReDim Preserve buf(1 To N)
        获取样式名数组 = buf
    End If
End Function

'——— 删除手工编号：根据“编号样式+编号格式”动态生成正则（仅匹配段首）
Public Function 生成删除编号规则集() As Variant
    Dim cfg As Variant, i As Long, pat As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    cfg = 获取所有级别参数()
    For i = 1 To UBound(cfg, 1)
        pat = 构造删除规则(CLng(cfg(i, 3)), CStr(cfg(i, 2)))
        If Len(pat) > 0 Then
            If Not dict.exists(pat) Then dict.Add pat, True
        End If
    Next i
    
    ' 加入中文序号：一、二、三、十一…（段首），后接可选标点与空格
    ' 说明：匹配 1~3 个中文数字（含 十/百/千 组合），后可有 、，.．。：: 等
    If Not dict.exists("^[ \t]*[一二三四五六七八九十百千]{1,3}\s*(?:[、,，:：．。.\-—–]\s*)?") Then
        dict.Add "^[ \t]*[一二三四五六七八九十百千]{1,3}\s*(?:[、,，:：．。.\-—–]\s*)?", True
    End If

    
    ' 加入通用兜底：裸数字+空白（含全角空格）
    If Not dict.exists("^\d+[ 　\t]+") Then dict.Add "^\d+[ 　\t]+", True
    
    If dict.Count = 0 Then
        生成删除编号规则集 = Array()
    Else
        生成删除编号规则集 = dict.Keys
    End If
End Function


'—— 删除手工编号：更宽容版本（段首匹配）
Private Function 构造删除规则(ByVal numStyle As Long, ByVal numFormat As String) As String
    Dim c As Long: c = 占位数(numFormat)
    Dim punct As String: punct = "[、,，:：．。.\-—–]"  ' 允许的后缀标点（可按需增减）

    ' 7 项：各类圈号数字（①/?/⑴ 等）
    If numStyle = wdListNumberStyleNumberInCircle Then
        构造删除规则 = "^[ \t]*[" & 构造项符号集() & "]\s*"
        Exit Function
    End If

    ' 6 款：（%n）或 (%n) —— 全/半角括号 + 允许空格
    If InStr(numFormat, "（%") > 0 Or InStr(numFormat, "(%") > 0 Then
        构造删除规则 = "^[ \t]*[（(]\s*\d+\s*[)）]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If

    ' 5 条：%n）或 %n) —— 数字 + 右括号（全/半角），允许空格
    If Right$(Trim$(numFormat), 1) = "）" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            构造删除规则 = "^[ \t]*\d+\s*[)）]\s*(?:" & punct & "\s*)?"
            Exit Function
        End If
    End If

    ' 多级：1.1 或 1． 1．1 —— 点前后可有空格，末尾点可选
    If c >= 2 Then
        构造删除规则 = "^[ \t]*\d+(?:\s*[\.．。]\s*\d+){1,}\s*(?:[\.．。])?\s*(?:" & punct & "\s*)?"
        Exit Function
    End If

    ' 单级：1 或 1. —— 但排除“1）/1)”（负向前瞻）
    If c = 1 Then
        构造删除规则 = "^[ \t]*\d+(?!\s*[)）])\s*(?:[\.．。]\s*)?(?:" & punct & "\s*)?"
        Exit Function
    End If

    ' 兜底：裸数字 + 空白
    构造删除规则 = "^[ \t]*\d+[ 　\t]+"
End Function
'====================（修复版）按优先级生成：款→条→项→4段→3段→2段→1段(带点)→1段(纯数字) ====================
Public Function 生成标题匹配规则集() As Variant
    Dim cfg As Variant, i As Long
    Dim sty As String, fmt As String, kind As Long
    Dim buckets(1 To 8) As Collection  ' （一）8类优先级桶
    Dim cat As Integer, pat As String
    Dim rowsCol As New Collection
    Dim rows() As Variant
    Dim p As Variant, j As Long, N As Long, k As Long
    Dim order As Variant
    
    '（二）初始化8个桶：1款 2条 3项 4四段 5三段 6二段 7单段带点 8单段纯数字
    For i = 1 To 8
        Set buckets(i) = New Collection
    Next i
    
    '（三）读取编号参数表
    cfg = 获取所有级别参数()
    
    '（四）把每条规则丢到对应优先级桶里
    For i = 1 To UBound(cfg, 1)
        sty = CStr(cfg(i, 1))
        fmt = CStr(cfg(i, 2))
        kind = CLng(cfg(i, 3))
        
        pat = 构造标题匹配规则(kind, fmt)
        If Len(pat) > 0 Then
            cat = 规则类别(kind, fmt)
            buckets(cat).Add Array(pat, sty)
        End If
    Next i
    
    '（五）按既定优先级拼接（更具体的在前）
    order = Array(1, 2, 3, 4, 5, 6, 7, 8)
    For Each p In order
        For j = 1 To buckets(p).Count
            rowsCol.Add buckets(p)(j)
        Next j
    Next p
    
    '（六）转成二维数组返回给调用方
    N = rowsCol.Count
    If N = 0 Then
        生成标题匹配规则集 = Array()
        Exit Function
    End If
    
    ReDim rows(1 To N, 1 To 2)
    For k = 1 To N
        rows(k, 1) = rowsCol(k)(0)  ' pattern
        rows(k, 2) = rowsCol(k)(1)  ' style
    Next k
    
    生成标题匹配规则集 = rows
End Function

'——（配套）把一种编号格式归类到 8 个优先级之一
Private Function 规则类别(ByVal numStyle As Long, ByVal numFmt As String) As Integer
    Dim c As Long: c = 占位数(numFmt)
    ' ① 项（圈号）→ 类别3
    If numStyle = wdListNumberStyleNumberInCircle Then
        规则类别 = 3: Exit Function
    End If
    ' ② 款（（n）/ (n)）→ 类别1
    If InStr(numFmt, "（%") > 0 Or InStr(numFmt, "(%") > 0 Then
        规则类别 = 1: Exit Function
    End If
    ' ③ 条（n）/ n)）→ 类别2
    If Right$(Trim$(numFmt), 1) = "）" Or Right$(Trim$(numFmt), 1) = ")" Then
        If InStr(numFmt, "%") > 0 Then 规则类别 = 2: Exit Function
    End If
    ' ④～⑦ 点分式标题
    Select Case c
        Case 4: 规则类别 = 4: Exit Function
        Case 3: 规则类别 = 5: Exit Function
        Case 2: 规则类别 = 6: Exit Function
        Case 1
            If InStr(numFmt, ".") > 0 Or InStr(numFmt, "．") > 0 Or InStr(numFmt, "。") > 0 Then
                规则类别 = 7     ' 单段带点，如“1.”
            Else
                规则类别 = 8     ' 单段纯数字，如“1”
            End If
        Case Else
            规则类别 = 8
    End Select
End Function

'====================（修复版）严格匹配：修正一级“纯数字”正则 ====================
Private Function 构造标题匹配规则( _
    ByVal numStyle As Long, _
    ByVal numFormat As String _
) As String
    
    Dim c As Long: c = 占位数(numFormat)
    Dim dot As String: dot = "[\.．。]"   ' 允许的点号（半/全角）
    
    '（一）项：只认圈号
    If numStyle = wdListNumberStyleNumberInCircle Then
        构造标题匹配规则 = "^[ \t]*[" & 构造项符号集() & "]\s*"
        Exit Function
    End If
    
    '（二）款：只认（n）/ (n)
    If InStr(numFormat, "（%") > 0 Or InStr(numFormat, "(%") > 0 Then
        构造标题匹配规则 = "^[ \t]*[（(]\s*\d+\s*[)）]\s*"
        Exit Function
    End If
    
    '（三）条：只认 n）/ n)
    If Right$(Trim$(numFormat), 1) = "）" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            构造标题匹配规则 = "^[ \t]*\d+\s*[)）]\s*"
            Exit Function
        End If
    End If
    
    '（四）标题 1~4：严格匹配“恰好 N 段”
    Select Case c
        Case 4
            构造标题匹配规则 = "^[ \t]*\d+\s*" & dot & "\s*\d+\s*" & dot & "\s*\d+\s*" & dot & "\s*\d+(?!\s*" & dot & "\s*\d)"
        Case 3
            构造标题匹配规则 = "^[ \t]*\d+\s*" & dot & "\s*\d+\s*" & dot & "\s*\d+(?!\s*" & dot & "\s*\d)"
        Case 2
            构造标题匹配规则 = "^[ \t]*\d+\s*" & dot & "\s*\d+(?!\s*" & dot & "\s*\d)"
        Case 1
            ' 你的一级格式是 “%1  ”（不带点，见参数表）→ 需排除 “1）/1)” 与 “1.1” 两种跟随
            '【关键修正】负向前瞻里用分组，一次排除两类前缀：右括号 或 “点+数字”
            If InStr(numFormat, ".") > 0 Or InStr(numFormat, "．") > 0 Or InStr(numFormat, "。") > 0 Then
                构造标题匹配规则 = "^[ \t]*\d+\s*" & dot & "(?!\s*\d)"
            Else
                构造标题匹配规则 = "^[ \t]*\d+(?!\s*(?:[)）]|" & dot & "\s*\d))"
            End If
        Case Else
            构造标题匹配规则 = ""
    End Select
End Function



'—— 公共：项用圈号数字集合（避免出现“？？？”）
Public Function 构造项符号集() As String
    Dim s As String, code As Long
    For code = &H2460 To &H2473: s = s & ChrW(code): Next code   ' ①..?
    For code = &H2474 To &H2487: s = s & ChrW(code): Next code   ' ⑴..⒇
    For code = &H2776 To &H277F: s = s & ChrW(code): Next code   ' ?..?
    For code = &H24EB To &H24F4: s = s & ChrW(code): Next code   ' ?..?
    For code = &H24F5 To &H24FE: s = s & ChrW(code): Next code   ' ?..?
    构造项符号集 = s
End Function


' 统计 "%n" 占位符数量
Private Function 占位数(ByVal fmt As String) As Long
    Dim i As Long, c As Long
    For i = 1 To Len(fmt)
        If mid$(fmt, i, 1) = "%" Then c = c + 1
    Next
    占位数 = c
End Function

'——— 读取“样式名”数组（针对指定文档；onlyExisting=True 时只返回该文档里存在的样式）
Public Function 获取样式名数组_针对文档(ByVal src As Document, Optional onlyExisting As Boolean = True) As Variant
    Dim cfg As Variant, i As Long, N As Long
    Dim buf() As String
    Dim sty As Style, name As String

    cfg = 获取所有级别参数()  ' 来自步骤（三）
    ReDim buf(1 To UBound(cfg, 1))
    N = 0

    For i = 1 To UBound(cfg, 1)
        name = CStr(cfg(i, 1))
        If onlyExisting Then
            On Error Resume Next
            Set sty = src.Styles(name)
            On Error GoTo 0
            If Not sty Is Nothing Then
                N = N + 1: buf(N) = name
                Set sty = Nothing
            End If
        Else
            N = N + 1: buf(N) = name
        End If
    Next i

    If N = 0 Then
        获取样式名数组_针对文档 = Array()
    Else
        ReDim Preserve buf(1 To N)
        获取样式名数组_针对文档 = buf
    End If
End Function

