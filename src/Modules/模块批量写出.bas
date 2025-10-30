Attribute VB_Name = "ģ������д��"
Option Explicit
'======================================================================
'  Word VBA | ������ǰ���̵� GitHub �Ѻýṹ������ʽ + README ֻ�����Զ�����
'  Entry : ����_��ǰ����_������ / ����_��ǰ����_ȫ��
'======================================================================

'============================��һ����������============================
'��Ҫ�󡿵�����Ŀ¼���Զ�������
Private Const BACKUP_ROOT As String = "E:\BaiduSyncdisk\Word-ToolBox-VBA"
'��Ҫ�󡿵���ǰ�Ƿ���� src Ŀ¼��True=��գ�������ļ�������
Private Const CLEAR_SRC_BEFORE_EXPORT As Boolean = True

' README �Զ�����ı�ǣ�ֻ�滻��Σ���������д���������ݣ�
Private Const MARK_BEGIN As String = "<!-- AUTO:EXPORT-BLOCK:BEGIN -->"
Private Const MARK_END   As String = "<!-- AUTO:EXPORT-BLOCK:END -->"

'��һ����ȷ���ư���������Сд�����У������ֻ���ڴ˴�׷�ӣ�
Private Function NamesToExport() As Variant
    NamesToExport = Array( _
        "frmDocObjectInspector", "ProgressForm", "��׼����ʽ������", _
        "ȫ�Ŀ��ж�ҳ����", "��ʽ_��׼��ҳ������", "��ʽ_������ʽһ������", "�Լ챨��", _
        "MOD_�����ʼ������", "MOD_���廽��", "MOD_�������ߺ���", "MOD_��ʽ����", _
        "���_1_������������ʽ", "���_2_�Զ�ƥ���������", "���_3_�༶ģ�嶨���ƥ��", _
        "���_4_�Զ�ɾ���ֹ����", "���_5_�Լ칤��", "���_6_�����ȵ�ȥ�����", _
        "���_���мӴ�", _
        "������_1_ͳһ�������ʽ", "������_1_ͳһ_�������ʽ", _
        "������_2_�Զ��������", "������_3_ȥ���˹����", "������_4_�˹���鹤��", _
        "������_5_Ԥ���", "������_6_�����Լ�", _
        "ȫ��_���_����Ӵ�", "ȫ��_���_���ж�ҳ", "ȫ��_���_ͳһ������ʽ", "ȫ��_���_��ȫ��ʽ��", _
        "ȫ��_ȫѡ��ǰ��", "ȫ��_��ʽ_���δʹ����ʽ", _
        "ͼƬ����_1_��ʽƥ��", "ͼƬ����_2_�Զ����", "ͼƬ����_3_ȥ���˹����", _
        "��ʽ_��׼�������ʽ", "��ʽ_ʩ����֯��ʽһ������", "��ʽ_�ļ�����_������", "��ʽ_����_ʩ����֯���" _
    )
End Function

'������ͨ���������������ǰ�رգ����ֿ����飻�Ժ���Ҫ�ټӣ�
Private Function PatternsToExport() As Variant
    PatternsToExport = Array()
End Function

'============================�������������============================
'��һ��ֻ������ǰ���̣���������
Public Sub ����_��ǰ����_������()
    ExportActiveProject False
End Sub

'������ֻ������ǰ���̣�ȫ�����
Public Sub ����_��ǰ����_ȫ��()
    ' ������������ȡCommit��Ϣ
    Dim commitMsg As String
    commitMsg = InputBox( _
        Prompt:="�����뱾�ε�����Commit��Ϣ������ʾ��README�У���" & vbCrLf & "���磺�޸����������߼�", _
        Title:="����Commit��Ϣ", _
        Default:="") ' Ĭ�Ͽ�ֵ
    
    ' �����û�ȡ��/δ��������
    If commitMsg = "" Then
        MsgBox "δ����Commit��Ϣ����ȡ������������", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    ' ���ú��ĵ������̣�������commit��Ϣ
    ExportActiveProject True, commitMsg

End Sub

'============================���������Ĺ���============================
Private Sub ExportActiveProject(ByVal exportAll As Boolean, Optional ByVal commitMsg As String = "")
    On Error GoTo FAIL

    '��һ���õ�ǰ����
    Dim proj As Object
    On Error Resume Next
    Set proj = Application.VBE.ActiveVBProject
    On Error GoTo 0
    If proj Is Nothing Then
        MsgBox "δ�ҵ�� VBA ���̡����ڹ��̴���ѡ��һ�����̺����ԡ�", vbExclamation, "����"
        Exit Sub
    End If
    If Not CanAccessProject(proj) Then
        MsgBox Join(Array( _
            "��ǰ���̱�������δ���á����ζ�VBA���̶���ģ�͵ķ��ʡ���", _
            "�뵽���ļ� �� ѡ�� �� �������� �� ������������ �� �����ã���ѡ���" _
        ), vbCrLf), vbCritical, "�޷����ʹ���"
        Exit Sub
    End If

    '������׼�� Git �Ѻ�Ŀ¼
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim root As String: root = EnsureFolder(BACKUP_ROOT, fso)
    Dim dirSrc As String: dirSrc = EnsureFolder(root & "\src", fso)
    Call EnsureFolder(dirSrc & "\Modules", fso)
    Call EnsureFolder(dirSrc & "\Classes", fso)
    Call EnsureFolder(dirSrc & "\Forms", fso)
    Call EnsureFolder(dirSrc & "\Documents", fso)
    Call EnsureFolder(root & "\manifest", fso)
    Call EnsureFolder(root & "\templates\Normal", fso)

    '��������ѡ���ȴ�
    Dim pf As Object
    On Error Resume Next
    Set pf = VBA.UserForms.Add("ProgressForm")
    On Error GoTo 0
    If Not pf Is Nothing Then
        pf.caption = "��������ǰ���̣���" & proj.name
        pf.Show vbModeless
        DoEvents
    End If

    '���ģ�ͳ������
    Dim total As Long: total = CountMatchesInProject(proj, exportAll)
    If total = 0 Then
        StatusPulse pf, "δƥ�䵽��Ҫ�����������"
        GoTo WRITE_META
    End If

    '���壩��վ��ļ����� src/* Ŀ¼��
    If CLEAR_SRC_BEFORE_EXPORT Then
        PurgeFolder dirSrc & "\Modules"
        PurgeFolder dirSrc & "\Classes"
        PurgeFolder dirSrc & "\Forms"
        PurgeFolder dirSrc & "\Documents"
        StatusPulse pf, "����վɵĵ����ļ���src/*����"
    End If

    '������������������ǣ�
    Dim done As Long: done = 0
    Dim comp As Object, log As String
    For Each comp In proj.VBComponents
        If exportAll Or IsTargetName(comp.name) Then
            Dim subDir As String, ext As String, dst As String, base As String
            subDir = SubFolderByType(comp.Type)
            ext = GuessExtByType(comp.Type)
            base = root & "\src\" & subDir & "\" & SafeFile(comp.name)
            dst = base & ext

            ' ���Ǿ��ļ����� .frx��
            If fso.FileExists(dst) Then On Error Resume Next: fso.DeleteFile dst, True: On Error GoTo 0
            If ext = ".frm" Then
                If fso.FileExists(base & ".frx") Then On Error Resume Next: fso.DeleteFile base & ".frx", True: On Error GoTo 0
            End If

            On Error Resume Next
            comp.Export dst
            If Err.Number = 0 Then
                log = log & "�� " & comp.name & "  ��  " & dst & vbCrLf
            Else
                log = log & "�� " & comp.name & "  ��  " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0

            done = done + 1
            UpdateBar pf, done, total, "������" & done & "/" & total & "���� " & comp.name
            If Not pf Is Nothing Then If pf.stopFlag Then Exit For
        End If
    Next

    '���ߣ�ͬʱ���� Normal.dotm�����ǣ�
    CopyNormalTemplate root & "\templates\Normal\normal.dotm"

WRITE_META:
    '���ˣ�д�����嵥 / ������־ / Git �ļ�
    WriteReferences proj, root & "\manifest\references.txt"
    WriteTextFile root & "\manifest\export_log.txt", log
    WriteGitignore root
    WriteReadme root, proj.name, commitMsg    ' �� ��������commit��Ϣ

    StatusPulse pf, "��ɡ���ƥ�䵽 " & CStr(total) & " ��������Ѵ��� " & CStr(done) & " ����"
    If Not pf Is Nothing Then Unload pf
    Application.StatusBar = False
    MsgBox "������ɣ���Ŀ¼��" & root, vbInformation, "����"
    Exit Sub

FAIL:
    If Not pf Is Nothing Then Unload pf
    Application.StatusBar = False
    MsgBox "���������з�������" & Err.Description, vbCritical, "����ʧ��"
End Sub

'============================���ģ��ж�/ͳ��/ӳ��============================
Private Function CanAccessProject(ByVal proj As Object) As Boolean
    On Error Resume Next
    Dim c As Long: c = proj.VBComponents.Count
    CanAccessProject = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function IsTargetName(ByVal compName As String) As Boolean
    Dim nm As Variant
    For Each nm In NamesToExport
        If UCase$(CStr(nm)) = UCase$(compName) Then IsTargetName = True: Exit Function
    Next
    For Each nm In PatternsToExport
        If compName Like CStr(nm) Then IsTargetName = True: Exit Function
    Next
End Function

Private Function CountMatchesInProject(ByVal proj As Object, ByVal exportAll As Boolean) As Long
    Dim comp As Object, N As Long
    If Not CanAccessProject(proj) Then Exit Function
    For Each comp In proj.VBComponents
        If exportAll Or IsTargetName(comp.name) Then N = N + 1
    Next
    CountMatchesInProject = N
End Function

Private Function SubFolderByType(ByVal vbCompType As Long) As String
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2
    Const vbext_ct_MSForm As Long = 3
    Const vbext_ct_Document As Long = 100
    Select Case vbCompType
        Case vbext_ct_StdModule:   SubFolderByType = "Modules"
        Case vbext_ct_ClassModule: SubFolderByType = "Classes"
        Case vbext_ct_MSForm:      SubFolderByType = "Forms"
        Case vbext_ct_Document:    SubFolderByType = "Documents"
        Case Else:                 SubFolderByType = "Modules"
    End Select
End Function

Private Function GuessExtByType(ByVal vbCompType As Long) As String
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2
    Const vbext_ct_MSForm As Long = 3
    Const vbext_ct_Document As Long = 100
    Select Case vbCompType
        Case vbext_ct_StdModule:   GuessExtByType = ".bas"
        Case vbext_ct_ClassModule: GuessExtByType = ".cls"
        Case vbext_ct_MSForm:      GuessExtByType = ".frm"   ' ����� .frx
        Case vbext_ct_Document:    GuessExtByType = ".cls"
        Case Else:                 GuessExtByType = ".bas"
    End Select
End Function

'============================���壩IO / Git / ״̬����============================
' ȷ��Ŀ¼����
' ȷ��Ŀ¼���ڣ��ݹ鴴����֧�ֶ༶/UNC/β��б�ܣ�
Private Function EnsureFolder(ByVal path As String, ByVal fso As Object) As String
    Dim p As String: p = path
    ' ȥ��ĩβ "\"������Ѹ�Ŀ¼��ɾ���մ����⣩
    If Len(p) > 3 And Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If p = "" Then EnsureFolder = "": Exit Function

    ' �Ѵ���ֱ�ӷ���
    If fso.FolderExists(p) Then
        EnsureFolder = p
        Exit Function
    End If

    ' ��ȷ���ϼ�Ŀ¼����
    Dim cut As Long: cut = InStrRev(p, "\")
    If cut > 0 Then
        Dim parent As String: parent = Left$(p, cut - 1)
        ' �����̷������� "C:"���� UNC ������ "\\server\share"��������
        If parent <> "" And Not fso.FolderExists(parent) Then
            ' ���� UNC��ȷ���� \\server\share Ϊֹ
            If Left$(p, 2) = "\\" Then
                Dim first As Long, second As Long
                first = InStr(3, p, "\")                   ' �������ַ�֮���ҵ�һ�� "\"
                If first > 0 Then second = InStr(first + 1, p, "\") ' ��������� "\"
                If second > 0 And Len(parent) >= second - 1 Then
                    ' parent ��Ȼ������һ�㣬�ݹ鼴��
                End If
            End If
            Call EnsureFolder(parent, fso)
        End If
    End If

    ' ������ǰĿ¼��������/Ȩ��ʧ���� MkDir ���ף�
    On Error Resume Next
    fso.CreateFolder p
    If Err.Number <> 0 Then
        Err.Clear
        MkDir p
    End If
    On Error GoTo 0

    EnsureFolder = p
End Function


' ���Ŀ¼���ļ�����ɾ��Ŀ¼��
Private Sub PurgeFolder(ByVal folderPath As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        Dim f As Object
        For Each f In fso.GetFolder(folderPath).Files
            f.Delete True
        Next
    End If
    On Error GoTo 0
End Sub

' д�ı���UTF-8 BOM�����ǣ�
Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2: .Charset = "utf-8": .Open
        .WriteText content
        .SaveToFile filePath, 2   ' 2=����
        .Close
    End With
End Sub

' ���ı���UTF-8 ���ȣ�ʧ���˻� ANSI��
Private Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo FAIL
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2: .Charset = "utf-8": .Open
        .LoadFromFile filePath
        ReadTextFile = .ReadText
        .Close
    End With
    Exit Function
FAIL:
    On Error GoTo 0
    If Len(Dir$(filePath)) > 0 Then
        Dim f As Integer: f = FreeFile
        Open filePath For Input As #f
        ReadTextFile = Input$(LOF(f), f)
        Close #f
    Else
        ReadTextFile = ""
    End If
End Function

' д�����嵥
Private Sub WriteReferences(ByVal proj As Object, ByVal outFile As String)
    On Error Resume Next
    Dim ref As Object, lines As Variant, s As String
    s = "Project: " & proj.name & vbCrLf & _
        "Exported: " & Format(Now, "yyyy-mm-dd HH:NN:SS") & vbCrLf & _
        String(40, "-") & vbCrLf
    For Each ref In proj.References
        s = s & ref.name & " | " & ref.GUID & " | v" & ref.Major & "." & ref.Minor & _
            " | " & ref.FullPath & vbCrLf
    Next
    On Error GoTo 0
    WriteTextFile outFile, s
End Sub

' д .gitignore�����ǣ������� Join �����м�����
Private Sub WriteGitignore(ByVal root As String)
    Dim lines As Variant
    lines = Array( _
        "# Office & VBA", _
        "~$*", _
        "*.tmp", _
        "*.lock", _
        "Thumbs.db", _
        ".DS_Store", _
        "" _
    )
    WriteTextFile root & "\.gitignore", Join(lines, vbCrLf)
End Sub

' д README�����״δ���ģ�壻�Ժ�������Զ����飩
Private Sub WriteReadme(ByVal root As String, ByVal projName As String, ByVal commitMsg As String)
    Dim path As String: path = root & "\README.md"
    Dim content As String, exists As Boolean
    exists = (Len(Dir$(path)) > 0)
    Dim autoBlock As String: autoBlock = BuildReadmeAutoBlock(projName, commitMsg)

    If Not exists Then
        ' ��һ�����ɣ�����ģ�� + �Զ�����
        WriteTextFile path, BuildReadmeTemplate(projName, autoBlock)
    Else
        ' �������У����滻�Զ����飬�����û��ֹ�����
        content = ReadTextFile(path)
        If InStr(1, content, MARK_BEGIN, vbTextCompare) > 0 And _
           InStr(1, content, MARK_END, vbTextCompare) > 0 Then
            content = ReplaceAutoBlock(content, MARK_BEGIN, MARK_END, autoBlock)
        Else
            content = content & vbCrLf & vbCrLf & autoBlock
        End If
        WriteTextFile path, content
    End If
End Sub
' ���� README ģ�壨���м�������
Private Function BuildReadmeTemplate(ByVal projName As String, ByVal autoBlock As String) As String
    Dim lines As Variant
    lines = Array( _
        "# " & projName & "��Word ��� / VBA ��Ŀ��", "", "���ֿ��� **��ǰ Word VBA ����** ��Դ�뵼�������ڰ汾������Э����", _
        "�����ű��Ḳ�Ǿɰ汾������ Normal.dotm һ�����ݡ�", "", "## Ŀ¼�ṹ", "- `src/Modules/`����׼ģ�飨.bas��", "- `src/Classes/`����ģ�飨.cls��", _
        "- `src/Forms/`�����壨.frm + .frx��", "- `src/Documents/`���ĵ�ģ�飨.cls��ThisDocument�ȣ�", "- `templates/Normal/normal.dotm`��Normal ģ�屸��", _
        "- `manifest/references.txt`�����������嵥", "- `manifest/export_log.txt`��������־", "", "## ʹ�÷���", "1. �� Word �� `Alt+F11` �� VBA �༭����", _
        "2. ���к꣺`����_��ǰ����_������`���� `����_��ǰ����_ȫ��`����", "3. ������Ŀ¼��`" & BACKUP_ROOT & "`��", "4. �״�ʹ�ã����� *�ļ� �� ѡ�� �� �������� �� ������������ �� ������* ��ѡ��**���ζ� VBA ���̶���ģ�͵ķ���**����", _
        "", "## �״����͵� GitHub��ʾ����", "```bash", "cd """ & BACKUP_ROOT & """", "git init", "git add .", "git commit -m ""init: ������ǰ����Դ��""", _
        "git branch -M main", "git remote add origin https://github.com/<your-account>/<repo>.git", "git push -u origin main", _
        "```", _
        "", _
        "## Լ��", _
        "- ģ��/����ע�Ͳ��ã�**һ**/**��**/������ţ��������ġ�", _
        "- ��������ǰ׺���壨�� `MOD_`��`������_`��`ȫ��_`����", _
        "- ��Ǩ�Ƶ� VSTO/VB.NET ʱ�������߼������ڶ���ģ���У�������ϡ�", _
        "", _
        autoBlock _
    )
    BuildReadmeTemplate = Join(lines, vbCrLf)
End Function

' ���� README �ġ��Զ����顱��ֻ�˶λᱻ���ǣ�
Private Function BuildReadmeAutoBlock(ByVal projName As String, ByVal commitMsg As String) As String
    Dim lines As Variant
    lines = Array( _
        MARK_BEGIN, _
        "### ������Ϣ���Զ����ɣ�", _
        "- �������� " & projName, _
        "- ����ʱ�䣺 " & Format(Now, "yyyy-mm-dd HH:NN:SS"), _
        "- ��Ŀ¼�� " & BACKUP_ROOT, _
        "- ���θ��£� " & commitMsg, _
        MARK_END _
    )
    BuildReadmeAutoBlock = Join(lines, vbCrLf)
End Function

' ���µ��Զ������滻������
Private Function ReplaceAutoBlock(ByVal content As String, _
                                  ByVal tagBegin As String, _
                                  ByVal tagEnd As String, _
                                  ByVal newBlock As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, content, tagBegin, vbTextCompare)
    p2 = InStr(p1 + Len(tagBegin), content, tagEnd, vbTextCompare)
    If p1 > 0 And p2 > p1 Then
        ReplaceAutoBlock = Left$(content, p1 - 1) & newBlock & mid$(content, p2 + Len(tagEnd))
    Else
        ReplaceAutoBlock = content & vbCrLf & vbCrLf & newBlock
    End If
End Function

' ���� Normal.dotm�����ǣ�
Private Sub CopyNormalTemplate(ByVal dstFullPath As String)
    On Error Resume Next
    Dim src As String: src = Application.NormalTemplate.FullName
    If Len(Dir$(src)) > 0 Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        EnsureFolder Left$(dstFullPath, InStrRev(dstFullPath, "\") - 1), fso
        fso.CopyFile src, dstFullPath, True
        StatusPulse Nothing, "�ѱ��� Normal.dotm �� " & dstFullPath
    End If
    On Error GoTo 0
End Sub

' ״̬��� / ����
Private Sub StatusPulse(ByVal pf As Object, ByVal msg As String)
    On Error Resume Next
    If Not pf Is Nothing Then
        pf.TextBoxStatus.text = pf.TextBoxStatus.text & vbCrLf & msg
        pf.TextBoxStatus.SelStart = Len(pf.TextBoxStatus.text)
        pf.TextBoxStatus.SelLength = 0
        pf.Repaint
    End If
    Application.StatusBar = msg
    DoEvents
    On Error GoTo 0
End Sub

Private Sub UpdateBar(ByVal pf As Object, ByVal cur As Long, ByVal total As Long, ByVal msg As String)
    On Error Resume Next
    Dim w As Integer: If total > 0 Then w = CInt(200 * (CDbl(cur) / total))
    If Not pf Is Nothing Then pf.UpdateProgressBar w, msg Else Application.StatusBar = msg
    On Error GoTo 0
End Sub

' �ļ�������
Private Function SafeFile(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace$(s, CStr(bad), "_")
    Next
    SafeFile = s
End Function

