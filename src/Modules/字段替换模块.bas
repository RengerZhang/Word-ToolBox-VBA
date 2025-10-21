Attribute VB_Name = "�ֶ��滻ģ��"
Option Explicit

'===========================================================
' �� Excel �������ɣ��� SaveAs �������� FileCopy��
' ���̣���ÿһ�� -> ��ǰ�ĵ� SaveAs ��Ŀ���ļ� -> �ڸø�����ȫ��Χ�滻
'     -> ����ر� -> ���´�ԭģ�壬������һ��
' ռλ����ʽ��{{��ͷ��}}���ı���/ҳüҳ��/������״ȫ�����滻
'===========================================================
Public Sub ��������_��������_SaveAs��()
    '���㣩������
    Const EXCEL_PATH As String = "C:\Users\Tony Zhang\Desktop\����\����.xlsx"
    Const SHEET_NAME As String = "Sheet1"
    Const OUTPUT_DIR As String = "C:\Users\Tony Zhang\Desktop\����\��������"
    Const L_DELIM As String = "{{"    ' ռλ�����
    Const R_DELIM As String = "}}"    ' ռλ���ҽ�
    Const FILENAME_PATTERN As String = "{{�������}}{{�ļ���}}.docx"

    Dim srcDoc As Document: Set srcDoc = ActiveDocument
    If Len(srcDoc.path) = 0 Then
        MsgBox "���Ȱѵ�ǰģ���ĵ����浽���̣�Ctrl+S���������С�", vbExclamation
        Exit Sub
    End If

    EnsureFolders OUTPUT_DIR     ' �ݹ鴴��Ŀ¼�������༶��

    '��һ���� Excel
    If Dir$(EXCEL_PATH) = "" Then
        MsgBox "�Ҳ��������ļ���" & EXCEL_PATH, vbExclamation: Exit Sub
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

    '���������д���
    For r = 2 To lastRow
        ' 1) �С��ֵ䣨�ֶ���=���б�ͷ���Զ���չ��
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

        ' 2) ��Ⱦ�ļ�����·��
        Dim outName As String, outPath As String
        outName = RenderPattern(FILENAME_PATTERN, dict, L_DELIM, R_DELIM)
        If Len(outName) = 0 Then outName = "��" & (r - 1) & "��.docx"
        outName = SanitizeFileName(outName)
        If LCase$(Right$(outName, 5)) <> ".docx" Then outName = outName & ".docx"
        outPath = CombinePath(OUTPUT_DIR, outName)

        ' ɾ��ͬ�������ļ������� SaveAs ����ֹ��
        On Error Resume Next
        If Dir$(outPath) <> "" Then
            SetAttr outPath, vbNormal
            Kill outPath
        End If
        On Error GoTo 0

        ' 3) �ѡ���ǰ�ĵ���ֱ�� SaveAs ��Ŀ���ļ�
        srcDoc.Save                              ' ȷ��ģ������
        srcDoc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument

        ' 4) ���±�������ĸ���������ȫ��Χ�滻��
        Call ReplaceByDict_Everywhere(ActiveDocument, dict, L_DELIM, R_DELIM)

        ' 5) ���沢�ر���ݳ�Ʒ
        ActiveDocument.Save
        ActiveDocument.Close SaveChanges:=False

        ' 6) ���´�ԭģ�壬������һ��
        Set srcDoc = Documents.Open(FileName:=srcPath, ReadOnly:=False, AddToRecentFiles:=False)

NextRow:
    Next r

    Application.ScreenUpdating = True
    MsgBox "��ɣ����������" & OUTPUT_DIR, vbInformation

CLEANUP:
    On Error Resume Next
    wb.Close SaveChanges:=False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing
End Sub

'==================== ȫ�����滻������/ҳü��/�ı���/����״�� ====================

Private Sub ReplaceByDict_Everywhere(ByVal doc As Document, ByVal d As Object, _
                                     ByVal LDelim As String, ByVal RDelim As String)
    Dim k As Variant, findText As String, rep As String, sec As Section, hf As HeaderFooter, shp As Shape
    For Each k In d.Keys
        findText = LDelim & CStr(k) & RDelim
        rep = NzStr(d(k))

        ' A. ���� Story���� wdTextFrameStory��
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

        ' B. ���ĵ�����״�����飬�ݹ飩
        For Each shp In doc.Shapes
            Replace_InShapeRecursive shp, findText, rep
        Next shp

        ' C. ҳü/ҳ���е���״
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

'==================== �������� ====================

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
        GetCellAsText = Format$(CDate(v), "yyyy��m��d��")
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


