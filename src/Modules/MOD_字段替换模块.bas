Attribute VB_Name = "MOD_字段替换模块"
Option Explicit

'===========================================================
'【用途】从 Excel 逐行读取数据 → 用 SaveAs 法批量生成 Word 文档
'【核心思路】
'（一）使用“当前打开的文档（模板）”作为版式来源；
'（二）对每一行：先把当前文档 SaveAs 成新文件（从而100%继承页眉页脚、文本框、分节等版式），
'      再在这份新文件中做“全范围替换”（正文/页眉/页脚/文本框/组内形状）；
'（三）保存并关闭新文件后，重新打开最初的模板，继续下一行。
'【占位符约定】模板中写 {{表头名}}，需与 Excel 首行表头一致。
'===========================================================
Public Sub 批量生成_塔吊方案_SaveAs法()
    '（一）参数区 —— 路径、工作表、占位符界定符、文件名模式
    Const EXCEL_PATH As String = "C:\Users\Tony Zhang\Desktop\测试\数据.xlsx"   '1）Excel 数据文件路径
    Const SHEET_NAME As String = "Sheet1"                                      '2）数据所在工作表名
    Const OUTPUT_DIR As String = "C:\Users\Tony Zhang\Desktop\测试\塔吊方案"     '3）输出目录
    Const L_DELIM As String = "{{"                                             '4）占位符左界，如 {{塔吊编号}}
    Const R_DELIM As String = "}}"                                             '5）占位符右界
    Const FILENAME_PATTERN As String = "{{塔吊编号}}{{文件名}}.docx"             '6）输出文件命名模板（可随时扩展）

    '（二）前置：确认模板状态 & 准备输出目录
    Dim srcDoc As Document: Set srcDoc = ActiveDocument         '1）当前文档作为“模板”使用
    If Len(srcDoc.path) = 0 Then                                '2）模板必须已保存到磁盘；否则 SaveAs/回开会失败
        MsgBox "请先把当前模板文档保存到磁盘（Ctrl+S）后再运行。", vbExclamation
        Exit Sub
    End If
    EnsureFolders OUTPUT_DIR                                    '3）确保输出目录（支持多级）存在

    '（三）打开 Excel（晚绑定，不需要设置引用）
    If Dir$(EXCEL_PATH) = "" Then                               '1）数据文件存在性检查
        MsgBox "找不到数据文件：" & EXCEL_PATH, vbExclamation: Exit Sub
    End If
    Dim xlApp As Object, wb As Object, ws As Object
    Set xlApp = CreateObject("Excel.Application")               '2）创建 Excel 进程（不可见）
    Set wb = xlApp.Workbooks.Open(EXCEL_PATH, ReadOnly:=True)   '3）只读打开数据工作簿
    Set ws = wb.Worksheets(SHEET_NAME)                          '4）定位到 Sheet1（可改）

    '（四）计算数据范围：最后一行/最后一列（首行是表头，第二行起是数据）
    Dim lastRow As Long, lastCol As Long, r As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(-4162).row         '1）xlUp：从 A 列底部向上找最后一行
    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column   '2）xlToLeft：从第1行右端向左找最后一列
    If lastRow < 2 Or lastCol < 1 Then GoTo CLEANUP             '3）无数据则直接收尾

    '（五）缓存模板磁盘路径；关闭屏幕刷新以提升效率
    Dim srcPath As String: srcPath = srcDoc.FullName            '1）模板的完整磁盘路径
    Application.ScreenUpdating = False                           '2）关刷新：生成过程更流畅

    '（六）主循环：逐行生成
    For r = 2 To lastRow
        ' 1）把“本行数据”读成字典（键=表头，值=单元格文本），自动支持新增字段
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim c As Long, key As String, val As String
        For c = 1 To lastCol
            key = Trim$(CStr(ws.Cells(1, c).Value))             'a）读取首行表头
            If Len(key) > 0 Then
                val = GetCellAsText(ws.Cells(r, c))             'b）读取该列在第 r 行的值（日期会格式化为“yyyy年m月d日”）
                dict(key) = val                                 'c）加入字典：dict("塔吊编号")="1#" 等
            End If
        Next c
        If dict.Count = 0 Then GoTo NextRow                     'd）空行保护

        ' 2）渲染文件名并得到输出路径
        Dim outName As String, outPath As String
        outName = RenderPattern(FILENAME_PATTERN, dict, L_DELIM, R_DELIM)   'a）按命名模板把占位符替换成值
        If Len(outName) = 0 Then outName = "第" & (r - 1) & "行.docx"       'b）兜底命名
        outName = SanitizeFileName(outName)                                  'c）清理非法文件名字符
        If LCase$(Right$(outName, 5)) <> ".docx" Then outName = outName & ".docx" 'd）确保扩展名
        outPath = CombinePath(OUTPUT_DIR, outName)                           'e）拼接成完整目标路径

        ' 3）清理同名旧文件（避免 SaveAs 被阻止）
        On Error Resume Next
        If Dir$(outPath) <> "" Then
            SetAttr outPath, vbNormal                                        'a）去掉只读等属性
            Kill outPath                                                     'b）删除旧文件
        End If
        On Error GoTo 0

        ' 4）关键步骤：将“当前模板文档”直接 SaveAs 成目标文件
        srcDoc.Save                                                          'a）先保存模板，确保磁盘版本最新
        srcDoc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument    'b）另存副本（确保版式100%继承）

        ' 5）在“刚保存出来的副本（即当前 ActiveDocument）”里做全覆盖替换
        Call ReplaceByDict_Everywhere(ActiveDocument, dict, L_DELIM, R_DELIM)

        ' 6）保存并关闭这份成品
        ActiveDocument.Save
        ActiveDocument.Close SaveChanges:=False

        ' 7）重新打开“最初的模板”进入下一轮（保持 srcDoc 指向模板）
        Set srcDoc = Documents.Open(FileName:=srcPath, ReadOnly:=False, AddToRecentFiles:=False)

NextRow:
    Next r

    '（七）收尾：恢复刷新 & 提示
    Application.ScreenUpdating = True
    MsgBox "完成，已输出到：" & OUTPUT_DIR, vbInformation

CLEANUP:
    '（八）Excel 资源释放（无论是否提前退出循环，都会走到这里）
    On Error Resume Next
    wb.Close SaveChanges:=False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing
End Sub

'==================== 全覆盖替换（正文/页眉脚/文本框/组形状） ====================
'【目的】Word 的文本不只在“正文故事”里：页眉/页脚、脚注、批注、文本框（TextFrame）、
'       以及“形状（包括组合形状）中的文字”都需要被覆盖替换。
'【做法】三步并行：
'   A. 遍历所有 StoryRanges（包含 wdTextFrameStory）；
'   B. 遍历主文档层 Shapes（对组合形状递归进入）；
'   C. 遍历各节的页眉/页脚的 Shapes（同样递归进入）。
Private Sub ReplaceByDict_Everywhere(ByVal doc As Document, ByVal d As Object, _
                                     ByVal LDelim As String, ByVal RDelim As String)
    Dim k As Variant, findText As String, rep As String, sec As Section, hf As HeaderFooter, shp As Shape
    For Each k In d.Keys
        findText = LDelim & CStr(k) & RDelim   '（一）本轮要查找的占位符文本，如 "{{塔吊编号}}"
        rep = NzStr(d(k))                      '（二）对应的替换值

        ' A. 所有 Story（含 wdTextFrameStory）：正文、页眉、页脚、脚注、批注、文本框等
        Dim rng As Range
        For Each rng In doc.StoryRanges
            Do
                With rng.Find
                    .ClearFormatting: .Replacement.ClearFormatting
                    .text = findText
                    .Replacement.text = rep
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchWildcards = False
                    .Execute Replace:=wdReplaceAll            '→ 把本 Story 中的占位符全部替换
                End With
                Set rng = rng.NextStoryRange                  '→ 跳到同类的下一个 Story（若有）
            Loop Until rng Is Nothing
        Next rng

        ' B. 主文档层形状（含组合）：处理 TextFrame.TextRange 内的文字
        For Each shp In doc.Shapes
            Replace_InShapeRecursive shp, findText, rep
        Next shp

        ' C. 页眉/页脚中的形状：不同于正文层的 doc.Shapes，这里要访问每个 HeaderFooter.Shapes
        For Each sec In doc.Sections
            For Each hf In sec.Headers
                For Each shp In hf.Shapes
                    Replace_InShapeRecursive shp, findText, rep
                Next shp
            Next hf
            For Each hf In sec.Footers
                For Each shp In hf.Shapes
                    Replace_InShapeRecursive shp, findText, rep
                Next shp
            Next hf
        Next sec
    Next k
End Sub

'【递归替换】进入 Shape：若是组合（msoGroup）则逐个子项；若是文本框（TextFrame.HasText）则对其 TextRange.Find
Private Sub Replace_InShapeRecursive(ByVal shp As Shape, ByVal findText As String, ByVal repText As String)
    On Error Resume Next                                    '（一）容错：部分形状可能不支持某些属性
    If shp.Type = msoGroup Then                             '（二）组合形状：递归进入子项
        Dim i As Long
        For i = 1 To shp.GroupItems.Count
            Replace_InShapeRecursive shp.GroupItems(i), findText, repText
        Next i
    Else                                                    '（三）普通形状：若含文本则执行查找替换
        If shp.TextFrame.HasText Then
            With shp.TextFrame.TextRange.Find
                .ClearFormatting: .Replacement.ClearFormatting
                .text = findText
                .Replacement.text = repText
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = False
                .Execute Replace:=wdReplaceAll
            End With
        End If
    End If
    On Error GoTo 0
End Sub

'==================== 基础工具 ====================

'（一）渲染命名模板：将模式串中的 "{{键}}" 全部替换成字典中的值
Private Function RenderPattern(ByVal pattern As String, ByVal d As Object, _
                               ByVal LDelim As String, ByVal RDelim As String) As String
    Dim k As Variant, s As String: s = pattern
    For Each k In d.Keys
        s = Replace$(s, LDelim & CStr(k) & RDelim, NzStr(d(k)))
    Next
    RenderPattern = s
End Function

'（二）统一把 Excel 单元格读取为“友好字符串”
'     1）如果是日期/时间，格式化为“yyyy年m月d日”；
'     2）否则转成去首尾空格的字符串。
Private Function GetCellAsText(ByVal cell As Object) As String
    Dim v: v = cell.Value
    If IsDate(v) Then
        GetCellAsText = Format$(CDate(v), "yyyy年m月d日")
    Else
        GetCellAsText = Trim$(CStr(v))
    End If
End Function

'（三）空值安全：把 Null/Empty 变成空串，避免替换时报错
Private Function NzStr(v) As String
    If IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

'（四）递归创建多级目录（"C:\a\b\c" 若 b/c 不存在会逐级创建）
Private Sub EnsureFolders(ByVal p As String)
    Dim parts() As String, i As Long, cur As String
    parts = Split(p, "\"): cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

'（五）简单路径拼接（兼容结尾是否带 “\”）
Private Function CombinePath(ByVal folder As String, ByVal name As String) As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        CombinePath = folder & name
    Else
        CombinePath = folder & "\" & name
    End If
End Function

'（六）清理非法文件名字符（\ / : * ? " < > |），并裁掉首尾空格
Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long: For i = LBound(bad) To UBound(bad): s = Replace$(s, bad(i), " "): Next
    SanitizeFileName = Trim$(s)
End Function


