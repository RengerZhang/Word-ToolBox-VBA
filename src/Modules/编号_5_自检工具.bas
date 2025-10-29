Attribute VB_Name = "编号_5_自检工具"
Option Explicit

'==========================================================
' 配置自检小工具（升级版）
' 变化点：
'   1) 报告文档：A4 横向，页边距 2cm
'   2) 全文默认字体：中文=宋体，西文/数字=Times New Roman，字号 10.5
'   3) 正则里的“项”符号不再手写，改用 Unicode 生成（支持多个集合）
'   4) 增加“匹配命中计数”列，快速定位 5~7 级没命中的原因
'
' 依赖：
'   - Public Function 获取所有级别参数()  ' 来自步骤（③）
'   - 若使用了 Mod配置中心，请保持“构造规则”的逻辑一致（建议同步替换）
'==========================================================
Sub 配置自检_生成报告()
    Dim srcDoc As Document   ' ← 要检查的源文档
    Dim rptDoc As Document   ' ← 新建的自检报告
    Dim cfg As Variant
    Dim i As Long, N As Long
    Dim rng As Range
    Dim tbl As Table
    Dim 行 As Long
    Dim 样式名 As String, 编号格式 As String
    Dim 编号样式值 As Long, 对齐cm As Single
    Dim 匹配正则 As String, 删除正则 As String
    Dim 样式是否存在 As String
    Dim 手工命中 As Long, 自动级数 As Long
    Dim 样式清单 As Variant, s As Variant
    Dim 删除全集 As Variant, p As Variant
    Dim 规则映射 As Variant, t As Long

    ' 1) 在新建报告之前，先牢牢抓住“源文档”
    Set srcDoc = ActiveDocument

    ' 2) 读取核心配置（来自步骤三）
    cfg = 获取所有级别参数()
    If IsEmpty(cfg) Then
        MsgBox "未读取到编号配置（获取所有级别参数() 返回为空）。", vbExclamation
        Exit Sub
    End If
    N = UBound(cfg, 1)

    ' 3) 新建报告文档（这会改变 ActiveDocument；所以后面一律用 srcDoc/rptDoc）
    Set rptDoc = Documents.Add
    With rptDoc.PageSetup
        .Orientation = wdOrientLandscape
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With
    With rptDoc.Styles(wdStyleNormal).Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 10.5
    End With
    rptDoc.content.Style = rptDoc.Styles(wdStyleNormal)

    ' 4) 标题
    Set rng = rptDoc.Range(0, 0)
    rng.text = "配置自检报告" & vbCrLf & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf

    ' 5) 表格（增加“自动编号计数”列，便于区分两类）
    Set tbl = rptDoc.Tables.Add(Range:=rng.Duplicate, NumRows:=N + 1, NumColumns:=10)
    With tbl
        .AllowAutoFit = True
        .AutoFitBehavior wdAutoFitWindow
        .rows(1).Range.bold = True
        .rows(1).Shading.BackgroundPatternColor = wdColorGray20
        .Borders.enable = True
        .Range.Font.NameFarEast = "宋体"
        .Range.Font.NameAscii = "Times New Roman"
        .Range.Font.NameOther = "Times New Roman"
        .Range.Font.Size = 10.5

        .cell(1, 1).Range.text = "级别"
        .cell(1, 2).Range.text = "样式名"
        .cell(1, 3).Range.text = "样式是否存在（源文档）"
        .cell(1, 4).Range.text = "编号格式"
        .cell(1, 5).Range.text = "编号样式"
        .cell(1, 6).Range.text = "对齐(cm)"
        .cell(1, 7).Range.text = "标题匹配正则（段首）"
        .cell(1, 8).Range.text = "删除编号正则（段首）"
        .cell(1, 9).Range.text = "命中计数（正则/手工）"
        .cell(1, 10).Range.text = "命中计数（自动编号=该级）"
        
        
        '——列宽比例预设（总宽 = 页面宽度 - 左右边距）
        Dim cw As Single
        Dim r(1 To 10) As Double
        Dim k As Long
        
        ' 关闭自适应，按固定宽度设置
        .AllowAutoFit = False
        
        ' 页面可用宽度（单位：pt）
        cw = rptDoc.PageSetup.PageWidth - rptDoc.PageSetup.LeftMargin - rptDoc.PageSetup.RightMargin
        
        ' 比例：1..10 列依次为
        ' 级别, 样式名, 是否存在, 编号格式, 编号样式, 对齐, 标题匹配正则, 删除编号正则, 段首前缀匹配数, 大纲级别匹配数
        r(1) = 0.06: r(2) = 0.12: r(3) = 0.1: r(4) = 0.12: r(5) = 0.1
        r(6) = 0.08: r(7) = 0.18: r(8) = 0.18: r(9) = 0.03: r(10) = 0.03
        
        ' 应用列宽（按点数设置）
        For k = 1 To 10
            .Columns(k).width = cw * r(k)
        Next k

    End With

    ' 6) 逐级填表（全部以 srcDoc 为准）
    For i = 1 To N
        样式名 = CStr(cfg(i, 1))
        编号格式 = CStr(cfg(i, 2))
        编号样式值 = CLng(cfg(i, 3))
        对齐cm = CSng(cfg(i, 4))

        样式是否存在 = IIf(样式存在(srcDoc, 样式名), "是", "否")

        匹配正则 = 构造标题匹配规则_自检v2(编号样式值, 编号格式)
        删除正则 = 构造删除规则_自检v2(编号样式值, 编号格式)

        手工命中 = 统计文档命中数(srcDoc, 匹配正则)
        自动级数 = 统计自动编号级数(srcDoc, i)

        tbl.cell(i + 1, 1).Range.text = CStr(i)
        tbl.cell(i + 1, 2).Range.text = 样式名
        tbl.cell(i + 1, 3).Range.text = 样式是否存在
        tbl.cell(i + 1, 4).Range.text = 编号格式
        tbl.cell(i + 1, 5).Range.text = 映射编号样式名(编号样式值)
        tbl.cell(i + 1, 6).Range.text = Format(对齐cm, "0.##")
        tbl.cell(i + 1, 7).Range.text = 匹配正则
        tbl.cell(i + 1, 8).Range.text = 删除正则
        tbl.cell(i + 1, 9).Range.text = CStr(手工命中)
        tbl.cell(i + 1, 10).Range.text = CStr(自动级数)
    Next i

    ' 7) 追加清单：一律以 srcDoc 为准
    Set rng = rptDoc.Range(rptDoc.content.End - 1, rptDoc.content.End - 1)
    
    '——在表格后追加“注释”说明两列指标
    Dim noteRng As Range
    Set noteRng = rptDoc.Range(tbl.Range.End, tbl.Range.End)
    noteRng.InsertParagraphAfter
    noteRng.Collapse wdCollapseEnd
    noteRng.text = _
        "注释：" & vbCrLf & _
        "? 段首前缀匹配数（按正则）：统计源文档中，该级对应的“段首编号形态”能在段落文本开头被正则匹配到的条数。" & vbCrLf & _
        "  仅看文本前缀；自动编号的数字不在段落文本里，因此不会计入此列。" & vbCrLf & _
        "? 大纲级别匹配数（按ListLevel）：统计源文档中，使用“多级大纲编号”且级别等于该级的段落数量。" & vbCrLf & _
        "  与段落文本无关，直接依据段落的 ListLevelNumber 判定。" & vbCrLf & _
        "理解方式（示例）：" & vbCrLf & _
        "  - 前缀匹配数≈0、级别匹配数>0：该级基本已用自动编号（正常）。" & vbCrLf & _
        "  - 前缀匹配数>0、级别匹配数=0：该级多为手工编号，建议先执行“标题匹配”与“自动多级编号”。" & vbCrLf & _
        "  - 两者都大：同级既有手工前缀又有自动编号，建议执行“去除手工编号”。" & vbCrLf & _
        "  - 两者都小：检查样式是否已创建/应用，或编号形态是否与配置一致。" & vbCrLf & vbCrLf

    
    rng.InsertAfter vbCrLf & "【A】目标样式名（源文档存在）" & vbCrLf
    样式清单 = 获取样式名数组_针对文档(srcDoc, True)
    If IsArray(样式清单) Then
        For Each s In 样式清单
            rng.InsertAfter " - " & CStr(s) & vbCrLf
        Next
    Else
        rng.InsertAfter "(无)" & vbCrLf
    End If

    rng.InsertAfter vbCrLf & "【B】删除手工编号规则集（段首）" & vbCrLf
    删除全集 = 生成删除编号规则集()
    If IsArray(删除全集) Then
        For Each p In 删除全集
            rng.InsertAfter " - " & CStr(p) & vbCrLf
        Next
    Else
        rng.InsertAfter "(无)" & vbCrLf
    End If

    rng.InsertAfter vbCrLf & "【C】标题匹配规则集（pattern → style）" & vbCrLf
    规则映射 = 生成标题匹配规则集()
    If IsArray(规则映射) Then
        For t = LBound(规则映射, 1) To UBound(规则映射, 1)
            rng.InsertAfter " - " & CStr(规则映射(t, 1)) & "  →  " & CStr(规则映射(t, 2)) & vbCrLf
        Next t
    Else
        rng.InsertAfter "(无)" & vbCrLf
    End If

    MsgBox "配置自检报告已生成（以源文档为口径）。", vbInformation
End Sub

' 统计“自动编号级别 == targetLevel”的段落数量
' 仅统计大纲编号（Outline Numbering），不把项目符号算进去
Private Function 统计自动编号级数(ByVal doc As Document, ByVal targetLevel As Long) As Long
    Dim p As Paragraph
    Dim c As Long, lvl As Long

    For Each p In doc.Paragraphs
        On Error Resume Next
        If p.Range.ListFormat.ListType = wdListOutlineNumbering Then   ' 只认多级大纲编号
            lvl = p.Range.ListFormat.ListLevelNumber
            If lvl = targetLevel Then c = c + 1
        End If
        On Error GoTo 0
    Next

    统计自动编号级数 = c
End Function

'—— 自检版：删除规则（与配置中心保持一致）
Private Function 构造删除规则_自检v2(ByVal numStyle As Long, ByVal numFormat As String) As String
    Dim c As Long: c = 统计占位数(numFormat)
    Dim punct As String: punct = "[、,，:：．。.\-—–]"

    If numStyle = wdListNumberStyleNumberInCircle Then
        构造删除规则_自检v2 = "^[ \t]*[" & 构造项符号集() & "]\s*"
        Exit Function
    End If
    If InStr(numFormat, "（%") > 0 Or InStr(numFormat, "(%") > 0 Then
        构造删除规则_自检v2 = "^[ \t]*[（(]\s*\d+\s*[)）]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If Right$(Trim$(numFormat), 1) = "）" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            构造删除规则_自检v2 = "^[ \t]*\d+\s*[)）]\s*(?:" & punct & "\s*)?"
            Exit Function
        End If
    End If
    If c >= 2 Then
        构造删除规则_自检v2 = "^[ \t]*\d+(?:\s*[\.．。]\s*\d+){1,}\s*(?:[\.．。])?\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If c = 1 Then
        构造删除规则_自检v2 = "^[ \t]*\d+(?!\s*[)）])\s*(?:[\.．。]\s*)?(?:" & punct & "\s*)?"
        Exit Function
    End If
    构造删除规则_自检v2 = "^[ \t]*\d+[ 　\t]+"
End Function


'—— 自检版：标题匹配规则（与配置中心保持一致）
Private Function 构造标题匹配规则_自检v2(ByVal numStyle As Long, ByVal numFormat As String) As String
    Dim c As Long: c = 统计占位数(numFormat)
    Dim punct As String: punct = "[、,，:：．。.\-—–]"

    If numStyle = wdListNumberStyleNumberInCircle Then
        构造标题匹配规则_自检v2 = "^[ \t]*[" & 构造项符号集() & "]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If InStr(numFormat, "（%") > 0 Or InStr(numFormat, "(%") > 0 Then
        构造标题匹配规则_自检v2 = "^[ \t]*[（(][ \t]*\d+[ \t]*[)）]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If Right$(Trim$(numFormat), 1) = "）" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            构造标题匹配规则_自检v2 = "^[ \t]*\d+[ \t]*[)）]\s*(?:" & punct & "\s*)?"
            Exit Function
        End If
    End If
    If c >= 2 Then
        构造标题匹配规则_自检v2 = "^[ \t]*\d+(?:\s*[\.．。]\s*\d+){1,}\s*(?:[\.．。])?\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If c = 1 Then
        构造标题匹配规则_自检v2 = "^[ \t]*\d+(?!\s*[)）])\s*(?:[\.．。]\s*)?(?:" & punct & "\s*)?"
        Exit Function
    End If
    构造标题匹配规则_自检v2 = "^[ \t]*\d+([ 　\t]+|$)"
End Function

