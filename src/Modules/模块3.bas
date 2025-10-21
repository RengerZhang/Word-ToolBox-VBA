Attribute VB_Name = "ģ��3"
Option Explicit

'==========================================================
' ��ϣ�Normal.dotm ����Ȩ��
' ���ã��ֱ��� Դ�ļ��Ƿ�ɶ���Ŀ���ļ����Ƿ��д������һ��ģ�⸴�Ʋ���
' ʹ�ã�������ճ����ģ�飬���� ���_Normal����Ȩ��()
' Ŀ��·����д����C:\Users\Tony Zhang\Desktop\VBAԴ���뱸��\
'==========================================================
Sub ���_Normal����Ȩ��()
    On Error GoTo ERR_HANDLER
    
    '��һ����λԴ��Ŀ��
    Dim src As String, srcDir As String
    Dim targetDir As String
    src = Application.NormalTemplate.FullName                     ' ��ǰʹ�õ� Normal.dotm
    srcDir = Left$(src, InStrRev(src, "\") - 1)                   ' Դ�ļ�����Ŀ¼
    targetDir = "C:\Users\Tony Zhang\Desktop\VBAԴ���뱸��\"      ' Ŀ�걸��Ŀ¼��д����
    
    '���������� Normal.dotm�����⡰δ���̡����¶�ʧ��
    Application.NormalTemplate.Save
    
    '������������
    Dim srcExists As Boolean, srcReadable As Boolean
    Dim tgtExists As Boolean, tgtWritable As Boolean
    Dim copyOk As Boolean, copyMsg As String
    Dim testDst As String, stamp As String
    
    srcExists = FileExists(src)
    If srcExists Then srcReadable = CanReadFile(src)
    
    tgtExists = FolderExists(targetDir)
    If tgtExists Then tgtWritable = HasWriteAccess(targetDir)
    
    '���ģ����ԡ�ģ�⸴�ơ���Ŀ��Ŀ¼����Ӱ�������ʽ���ݣ����ƺ�����ɾ����
    If srcReadable And tgtWritable Then
        stamp = Format(Now, "yyyymmdd_HHNNSS")
        testDst = AddSlash(targetDir) & "Normal_permtest_" & stamp & ".dotm"
        On Error Resume Next
        FileCopy src, testDst
        If Err.Number = 0 Then
            copyOk = True
            ' ���Ƴɹ�����ɾ�������ļ�
            Kill testDst
        Else
            copyOk = False
            copyMsg = "ģ�⸴��ʧ�ܣ�" & Err.Number & "���� " & Err.Description
        End If
        On Error GoTo ERR_HANDLER
    End If
    
    '���壩���ܱ���
    Dim markOK As String, markNG As String
    markOK = "?"
    markNG = "?"
    
    Dim REPORT As String
    REPORT = ""
    REPORT = REPORT & "��Դ�ļ���" & vbCrLf
    REPORT = REPORT & "·���� " & src & vbCrLf
    REPORT = REPORT & "�����ļ��У� " & srcDir & vbCrLf
    REPORT = REPORT & "���ڣ� " & IIf(srcExists, markOK, markNG) & vbCrLf
    REPORT = REPORT & "�ɶ��� " & IIf(srcReadable, markOK, markNG) & vbCrLf & vbCrLf
    
    REPORT = REPORT & "��Ŀ���ļ��С�" & vbCrLf
    REPORT = REPORT & "·���� " & targetDir & vbCrLf
    REPORT = REPORT & "���ڣ� " & IIf(tgtExists, markOK, markNG) & vbCrLf
    REPORT = REPORT & "��д�� " & IIf(tgtWritable, markOK, markNG) & vbCrLf & vbCrLf
    
    REPORT = REPORT & "��ģ�⸴�ơ�" & vbCrLf
    If srcReadable And tgtWritable Then
        REPORT = REPORT & "����� " & IIf(copyOk, "�ɹ� " & markOK, "ʧ�� " & markNG) & vbCrLf
        If Not copyOk Then REPORT = REPORT & copyMsg & vbCrLf
    Else
        REPORT = REPORT & "δִ�У���Դ���ɶ���Ŀ�겻��д��" & vbCrLf
    End If
    
    '�����������жϽ��ۣ����ж��߼���
    REPORT = REPORT & vbCrLf & "�����ۡ�" & vbCrLf
    If Not srcReadable And Not tgtWritable Then
        REPORT = REPORT & "Դ�ļ����ɶ� + Ŀ��Ŀ¼����д�����߾�����Ȩ��/�������⡣" & vbCrLf
    ElseIf Not srcReadable Then
        REPORT = REPORT & "Դ�ļ���Normal.dotm ���������ļ��У��޶�ȡȨ�޻�ռ�á�" & vbCrLf
    ElseIf Not tgtWritable Then
        REPORT = REPORT & "Ŀ���ļ�����д��Ȩ�ޣ���ϵͳ��ȫ�������أ���" & vbCrLf
    ElseIf srcReadable And tgtWritable And Not copyOk Then
        REPORT = REPORT & "��/д���ͨ������������ʧ�ܣ�����ǰ�ȫ���ԣ��硰�ܿ��ļ��з��ʡ�����ȫ������ء�" & vbCrLf
    Else
        REPORT = REPORT & "Դ�ɶ���Ŀ���д��ģ�⸴����������ǰʧ�ܿ���Ϊ��ʱռ�û�·�����⡣" & vbCrLf
    End If
    
    MsgBox REPORT, IIf(copyOk, vbInformation, vbExclamation), "Normal.dotm ����Ȩ�����"
    Exit Sub

ERR_HANDLER:
    MsgBox "��Ϲ��̳���" & Err.Number & "���� " & Err.Description, vbCritical, "����"
End Sub

'------------------------------------------
'������1���ļ��Ƿ����
'------------------------------------------
Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir$(filePath, vbNormal Or vbReadOnly Or vbSystem Or vbHidden)) > 0)
End Function

'------------------------------------------
'������2���ļ����Ƿ����
'------------------------------------------
Private Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then FolderExists = False: Exit Function
    FolderExists = (Len(Dir$(AddSlash(folderPath), vbDirectory)) > 0)
End Function

'------------------------------------------
'������3�������ļ��ɶ������Դ�Ϊֻ��
'------------------------------------------
Private Function CanReadFile(ByVal filePath As String) As Boolean
    On Error GoTo NG
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary Access Read As #ff
    Close #ff
    CanReadFile = True
    Exit Function
NG:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    CanReadFile = False
End Function

'------------------------------------------
'������4������Ŀ¼��д�����Դ�����ɾ����ʱ�ļ�
'------------------------------------------
Private Function HasWriteAccess(ByVal folderPath As String) As Boolean
    On Error GoTo NG
    Dim testFile As String, ff As Integer
    folderPath = AddSlash(folderPath)
    testFile = folderPath & "~write_test_" & Format(Now, "yyyymmdd_HHNNSS") & ".tmp"
    ff = FreeFile
    Open testFile For Output As #ff
    Print #ff, "test"
    Close #ff
    Kill testFile
    HasWriteAccess = True
    Exit Function
NG:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    HasWriteAccess = False
End Function

'------------------------------------------
'������5��ͳһ��ȫĿ¼��б��
'------------------------------------------
Private Function AddSlash(ByVal p As String) As String
    If Len(p) = 0 Then
        AddSlash = ""
    ElseIf Right$(p, 1) = "\" Then
        AddSlash = p
    Else
        AddSlash = p & "\"
    End If
End Function

