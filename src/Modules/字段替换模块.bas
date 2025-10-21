Attribute VB_Name = "字段替换模块"
Option Explicit

'===========================================================
' 从 Excel 逐行生成：用 SaveAs 法（不再 FileCopy）
' 流程：对每一行 -> 当前文档 SaveAs 成目标文件 -> 在该副本里全范围替换
'     -> 保存关闭 -> 重新打开原模板，进入下一行
' 占位符格式：{{表头名}}；文本框/页眉页脚/组内形状全部会替换
'===========================================================
Public Sub 批量生成_塔吊方案_SaveAs法()
    '（零）参数区
    Const EXCEL_PATH As String = "C:\Users\Tony Zhang\Desktop\测试\数据.xlsx"
    Const SHEET_NAME As String = "Sheet1"
    Const OUTPUT_DIR As String = "C:\Users\Tony Zhang\Desktop\测试\塔吊方案"
    Const L_DELIM As String = "{{"    ' 占位符左界
    Const R_DELIM As String = "}}"    ' 占位符右界
    Const FILENAME_PATTERN As String = "{{塔吊编号}}{{文件名}}.docx"

    Dim srcDoc As Document: Set srcDoc = ActiveDocument
    If Len(srcDoc.path) = 0 Then
        MsgBox "请先把当前模板文档保存到磁盘（Ctrl+S）后再运行。", vbExclamation
        Exit Sub
    End If

    EnsureFolders OUTPUT_DIR     ' 递归创建目录（包含多级）

    '（一）打开 Excel
    If Dir$(EXCEL_PATH) = "" Then
        MsgBox "找不到数据文件：" & EXCEL_PATH, vbExclamation: Exit Sub
    End If
    Dim xlApp As Object, wb As Object, ws As Object
    Set xlApp = CreateObject("Excel.Application")
    Set wb = xlApp.Workbooks.Open(EXCEL_PATH, ReadOnly:=True)
    Set ws = wb.Worksheets(SHEET_NAME)

    Dim lastRow As Long, lastCol As Long, r As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(-4162).row       'xlUp
    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column  'xlToLeft
    If lastRow < 2 Or lastCol < 1 Then GoTo CLEANUP

    Dim srcPath As String: srcPath = srcDoc.FullName
    Application.ScreenUpdating = False

    '（二）逐行处理
    For r = 2 To lastRow
        ' 1) 行→字典（字段名=首行表头，自动扩展）
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim c As Long, key As String, val As String
        For c = 1 To lastCol
            key = Trim$(CStr(ws.Cells(1, c).Value))
            If Len(key) > 0 Then
                val = GetCellAsText(ws.Cells(r, c))
                dict(key) = val
            End If
        Next c
        If dict.Count = 0 Then GoTo NextRow

        ' 2) 渲染文件名与路径
        Dim outName As String, outPath As String
        outName = RenderPattern(FILENAME_PATTERN, dict, L_DELIM, R_DELIM)
        If Len(outName) = 0 Then outName = "第" & (r - 1) & "行.docx"
        outName = SanitizeFileName(outName)
        If LCase$(Right$(outName, 5)) <> ".docx" Then outName = outName & ".docx"
        outPath = CombinePath(OUTPUT_DIR, outName)

        ' 删除同名现有文件（避免 SaveAs 被阻止）
        On Error Resume Next
        If Dir$(outPath) <> "" Then
            SetAttr outPath, vbNormal
            Kill outPath
        End If
        On Error GoTo 0

        ' 3) 把“当前文档”直接 SaveAs 成目标文件
        srcDoc.Save                              ' 确保模板落盘
        srcDoc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument

        ' 4) 在新保存出来的副本中做【全范围替换】
        Call ReplaceByDict_Everywhere(ActiveDocument, dict, L_DELIM, R_DELIM)

        ' 5) 保存并关闭这份成品
        ActiveDocument.Save
        ActiveDocument.Close SaveChanges:=False

        ' 6) 重新打开原模板，继续下一行
        Set srcDoc = Documents.Open(FileName:=srcPath, ReadOnly:=False, AddToRecentFiles:=False)

NextRow:
    Next r

    Application.ScreenUpdating = True
    MsgBox "完成，已输出到：" & OUTPUT_DIR, vbInformation

CLEANUP:
    On Error Resume Next
    wb.Close SaveChanges:=False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing
End Sub

'==================== 全覆盖替换（正文/页眉脚/文本框/组形状） ====================

Private Sub ReplaceByDict_Everywhere(ByVal doc As Document, ByVal d As Object, _
                                     ByVal LDelim As String, ByVal RDelim As String)
    Dim k As Variant, findText As String, rep As String, sec As Section, hf As HeaderFooter, shp As Shape
    For Each k In d.Keys
        findText = LDelim & CStr(k) & RDelim
        rep = NzStr(d(k))

        ' A. 所有 Story（含 wdTextFrameStory）
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
                    .Execute Replace:=wdReplaceAll
                End With
                Set rng = rng.NextStoryRange
            Loop Until rng Is Nothing
        Next rng

        ' B. 主文档层形状（含组，递归）
        For Each shp In doc.Shapes
            Replace_InShapeRecursive shp, findText, rep
        Next shp

        ' C. 页眉/页脚中的形状
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

Private Sub Replace_InShapeRecursive(ByVal shp As Shape, ByVal findText As String, ByVal repText As String)
    On Error Resume Next
    If shp.Type = msoGroup Then
        Dim i As Long
        For i = 1 To shp.GroupItems.Count
            Replace_InShapeRecursive shp.GroupItems(i), findText, repText
        Next i
    Else
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

Private Function RenderPattern(ByVal pattern As String, ByVal d As Object, _
                               ByVal LDelim As String, ByVal RDelim As String) As String
    Dim k As Variant, s As String: s = pattern
    For Each k In d.Keys
        s = Replace$(s, LDelim & CStr(k) & RDelim, NzStr(d(k)))
    Next
    RenderPattern = s
End Function

Private Function GetCellAsText(ByVal cell As Object) As String
    Dim v: v = cell.Value
    If IsDate(v) Then
        GetCellAsText = Format$(CDate(v), "yyyy年m月d日")
    Else
        GetCellAsText = Trim$(CStr(v))
    End If
End Function

Private Function NzStr(v) As String
    If IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Sub EnsureFolders(ByVal p As String)
    Dim parts() As String, i As Long, cur As String
    parts = Split(p, "\"): cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

Private Function CombinePath(ByVal folder As String, ByVal name As String) As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        CombinePath = folder & name
    Else
        CombinePath = folder & "\" & name
    End If
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long: For i = LBound(bad) To UBound(bad): s = Replace$(s, bad(i), " "): Next
    SanitizeFileName = Trim$(s)
End Function


