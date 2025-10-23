Attribute VB_Name = "MOD_�ֶ��滻ģ��"
Option Explicit

'===========================================================
'����;���� Excel ���ж�ȡ���� �� �� SaveAs ���������� Word �ĵ�
'������˼·��
'��һ��ʹ�á���ǰ�򿪵��ĵ���ģ�壩����Ϊ��ʽ��Դ��
'��������ÿһ�У��Ȱѵ�ǰ�ĵ� SaveAs �����ļ����Ӷ�100%�̳�ҳüҳ�š��ı��򡢷ֽڵȰ�ʽ����
'      ����������ļ�������ȫ��Χ�滻��������/ҳü/ҳ��/�ı���/������״����
'���������沢�ر����ļ������´������ģ�壬������һ�С�
'��ռλ��Լ����ģ����д {{��ͷ��}}������ Excel ���б�ͷһ�¡�
'===========================================================
Public Sub ��������_��������_SaveAs��()
    '��һ�������� ���� ·����������ռλ���綨�����ļ���ģʽ
    Const EXCEL_PATH As String = "C:\Users\Tony Zhang\Desktop\����\����.xlsx"   '1��Excel �����ļ�·��
    Const SHEET_NAME As String = "Sheet1"                                      '2���������ڹ�������
    Const OUTPUT_DIR As String = "C:\Users\Tony Zhang\Desktop\����\��������"     '3�����Ŀ¼
    Const L_DELIM As String = "{{"                                             '4��ռλ����磬�� {{�������}}
    Const R_DELIM As String = "}}"                                             '5��ռλ���ҽ�
    Const FILENAME_PATTERN As String = "{{�������}}{{�ļ���}}.docx"             '6������ļ�����ģ�壨����ʱ��չ��

    '������ǰ�ã�ȷ��ģ��״̬ & ׼�����Ŀ¼
    Dim srcDoc As Document: Set srcDoc = ActiveDocument         '1����ǰ�ĵ���Ϊ��ģ�塱ʹ��
    If Len(srcDoc.path) = 0 Then                                '2��ģ������ѱ��浽���̣����� SaveAs/�ؿ���ʧ��
        MsgBox "���Ȱѵ�ǰģ���ĵ����浽���̣�Ctrl+S���������С�", vbExclamation
        Exit Sub
    End If
    EnsureFolders OUTPUT_DIR                                    '3��ȷ�����Ŀ¼��֧�ֶ༶������

    '�������� Excel����󶨣�����Ҫ�������ã�
    If Dir$(EXCEL_PATH) = "" Then                               '1�������ļ������Լ��
        MsgBox "�Ҳ��������ļ���" & EXCEL_PATH, vbExclamation: Exit Sub
    End If
    Dim xlApp As Object, wb As Object, ws As Object
    Set xlApp = CreateObject("Excel.Application")               '2������ Excel ���̣����ɼ���
    Set wb = xlApp.Workbooks.Open(EXCEL_PATH, ReadOnly:=True)   '3��ֻ�������ݹ�����
    Set ws = wb.Worksheets(SHEET_NAME)                          '4����λ�� Sheet1���ɸģ�

    '���ģ��������ݷ�Χ�����һ��/���һ�У������Ǳ�ͷ���ڶ����������ݣ�
    Dim lastRow As Long, lastCol As Long, r As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(-4162).row         '1��xlUp���� A �еײ����������һ��
    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column   '2��xlToLeft���ӵ�1���Ҷ����������һ��
    If lastRow < 2 Or lastCol < 1 Then GoTo CLEANUP             '3����������ֱ����β

    '���壩����ģ�����·�����ر���Ļˢ��������Ч��
    Dim srcPath As String: srcPath = srcDoc.FullName            '1��ģ�����������·��
    Application.ScreenUpdating = False                           '2����ˢ�£����ɹ��̸�����

    '��������ѭ������������
    For r = 2 To lastRow
        ' 1���ѡ��������ݡ������ֵ䣨��=��ͷ��ֵ=��Ԫ���ı������Զ�֧�������ֶ�
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim c As Long, key As String, val As String
        For c = 1 To lastCol
            key = Trim$(CStr(ws.Cells(1, c).Value))             'a����ȡ���б�ͷ
            If Len(key) > 0 Then
                val = GetCellAsText(ws.Cells(r, c))             'b����ȡ�����ڵ� r �е�ֵ�����ڻ��ʽ��Ϊ��yyyy��m��d�ա���
                dict(key) = val                                 'c�������ֵ䣺dict("�������")="1#" ��
            End If
        Next c
        If dict.Count = 0 Then GoTo NextRow                     'd�����б���

        ' 2����Ⱦ�ļ������õ����·��
        Dim outName As String, outPath As String
        outName = RenderPattern(FILENAME_PATTERN, dict, L_DELIM, R_DELIM)   'a��������ģ���ռλ���滻��ֵ
        If Len(outName) = 0 Then outName = "��" & (r - 1) & "��.docx"       'b����������
        outName = SanitizeFileName(outName)                                  'c������Ƿ��ļ����ַ�
        If LCase$(Right$(outName, 5)) <> ".docx" Then outName = outName & ".docx" 'd��ȷ����չ��
        outPath = CombinePath(OUTPUT_DIR, outName)                           'e��ƴ�ӳ�����Ŀ��·��

        ' 3������ͬ�����ļ������� SaveAs ����ֹ��
        On Error Resume Next
        If Dir$(outPath) <> "" Then
            SetAttr outPath, vbNormal                                        'a��ȥ��ֻ��������
            Kill outPath                                                     'b��ɾ�����ļ�
        End If
        On Error GoTo 0

        ' 4���ؼ����裺������ǰģ���ĵ���ֱ�� SaveAs ��Ŀ���ļ�
        srcDoc.Save                                                          'a���ȱ���ģ�壬ȷ�����̰汾����
        srcDoc.SaveAs2 FileName:=outPath, FileFormat:=wdFormatXMLDocument    'b����渱����ȷ����ʽ100%�̳У�

        ' 5���ڡ��ձ�������ĸ���������ǰ ActiveDocument��������ȫ�����滻
        Call ReplaceByDict_Everywhere(ActiveDocument, dict, L_DELIM, R_DELIM)

        ' 6�����沢�ر���ݳ�Ʒ
        ActiveDocument.Save
        ActiveDocument.Close SaveChanges:=False

        ' 7�����´򿪡������ģ�塱������һ�֣����� srcDoc ָ��ģ�壩
        Set srcDoc = Documents.Open(FileName:=srcPath, ReadOnly:=False, AddToRecentFiles:=False)

NextRow:
    Next r

    '���ߣ���β���ָ�ˢ�� & ��ʾ
    Application.ScreenUpdating = True
    MsgBox "��ɣ����������" & OUTPUT_DIR, vbInformation

CLEANUP:
    '���ˣ�Excel ��Դ�ͷţ������Ƿ���ǰ�˳�ѭ���������ߵ����
    On Error Resume Next
    wb.Close SaveChanges:=False
    xlApp.Quit
    Set ws = Nothing: Set wb = Nothing: Set xlApp = Nothing
End Sub

'==================== ȫ�����滻������/ҳü��/�ı���/����״�� ====================
'��Ŀ�ġ�Word ���ı���ֻ�ڡ����Ĺ��¡��ҳü/ҳ�š���ע����ע���ı���TextFrame����
'       �Լ�����״�����������״���е����֡�����Ҫ�������滻��
'���������������У�
'   A. �������� StoryRanges������ wdTextFrameStory����
'   B. �������ĵ��� Shapes���������״�ݹ���룩��
'   C. �������ڵ�ҳü/ҳ�ŵ� Shapes��ͬ���ݹ���룩��
Private Sub ReplaceByDict_Everywhere(ByVal doc As Document, ByVal d As Object, _
                                     ByVal LDelim As String, ByVal RDelim As String)
    Dim k As Variant, findText As String, rep As String, sec As Section, hf As HeaderFooter, shp As Shape
    For Each k In d.Keys
        findText = LDelim & CStr(k) & RDelim   '��һ������Ҫ���ҵ�ռλ���ı����� "{{�������}}"
        rep = NzStr(d(k))                      '��������Ӧ���滻ֵ

        ' A. ���� Story���� wdTextFrameStory�������ġ�ҳü��ҳ�š���ע����ע���ı����
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
                    .Execute Replace:=wdReplaceAll            '�� �ѱ� Story �е�ռλ��ȫ���滻
                End With
                Set rng = rng.NextStoryRange                  '�� ����ͬ�����һ�� Story�����У�
            Loop Until rng Is Nothing
        Next rng

        ' B. ���ĵ�����״������ϣ������� TextFrame.TextRange �ڵ�����
        For Each shp In doc.Shapes
            Replace_InShapeRecursive shp, findText, rep
        Next shp

        ' C. ҳü/ҳ���е���״����ͬ�����Ĳ�� doc.Shapes������Ҫ����ÿ�� HeaderFooter.Shapes
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

'���ݹ��滻������ Shape��������ϣ�msoGroup���������������ı���TextFrame.HasText������� TextRange.Find
Private Sub Replace_InShapeRecursive(ByVal shp As Shape, ByVal findText As String, ByVal repText As String)
    On Error Resume Next                                    '��һ���ݴ�������״���ܲ�֧��ĳЩ����
    If shp.Type = msoGroup Then                             '�����������״���ݹ��������
        Dim i As Long
        For i = 1 To shp.GroupItems.Count
            Replace_InShapeRecursive shp.GroupItems(i), findText, repText
        Next i
    Else                                                    '��������ͨ��״�������ı���ִ�в����滻
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

'��һ����Ⱦ����ģ�壺��ģʽ���е� "{{��}}" ȫ���滻���ֵ��е�ֵ
Private Function RenderPattern(ByVal pattern As String, ByVal d As Object, _
                               ByVal LDelim As String, ByVal RDelim As String) As String
    Dim k As Variant, s As String: s = pattern
    For Each k In d.Keys
        s = Replace$(s, LDelim & CStr(k) & RDelim, NzStr(d(k)))
    Next
    RenderPattern = s
End Function

'������ͳһ�� Excel ��Ԫ���ȡΪ���Ѻ��ַ�����
'     1�����������/ʱ�䣬��ʽ��Ϊ��yyyy��m��d�ա���
'     2������ת��ȥ��β�ո���ַ�����
Private Function GetCellAsText(ByVal cell As Object) As String
    Dim v: v = cell.Value
    If IsDate(v) Then
        GetCellAsText = Format$(CDate(v), "yyyy��m��d��")
    Else
        GetCellAsText = Trim$(CStr(v))
    End If
End Function

'��������ֵ��ȫ���� Null/Empty ��ɿմ��������滻ʱ����
Private Function NzStr(v) As String
    If IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

'���ģ��ݹ鴴���༶Ŀ¼��"C:\a\b\c" �� b/c �����ڻ��𼶴�����
Private Sub EnsureFolders(ByVal p As String)
    Dim parts() As String, i As Long, cur As String
    parts = Split(p, "\"): cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

'���壩��·��ƴ�ӣ����ݽ�β�Ƿ�� ��\����
Private Function CombinePath(ByVal folder As String, ByVal name As String) As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        CombinePath = folder & name
    Else
        CombinePath = folder & "\" & name
    End If
End Function

'����������Ƿ��ļ����ַ���\ / : * ? " < > |�������õ���β�ո�
Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long: For i = LBound(bad) To UBound(bad): s = Replace$(s, bad(i), " "): Next
    SanitizeFileName = Trim$(s)
End Function


