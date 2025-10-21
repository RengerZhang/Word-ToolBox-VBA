Attribute VB_Name = "ģ��1"
Option Explicit

'==========================================================
' һ����� Normal.dotm �е�ָ����������� .dotm
' ʹ�÷�����
' 1���޸� �ٴ��ġ�Ҫ���������嵥��
' 2������ Pack_Normal_Modules_ToDotm
' 3���� �ڴ��޸����·����Ĭ�ϵ����棩
'==========================================================
Sub Pack_Normal_Modules_ToDotm()
    On Error GoTo ERRH

    '��һ���� ����Ҫ���������嵥�����Ʊ����� VBE ��һ�£�
    Dim modNames As Variant, clsNames As Variant, frmNames As Variant
    modNames = Array( _
        "modCommon", _               ' ʾ������������
        "modTitleMatch", _           ' ʾ��������ƥ��
        "modAutoNumber", _           ' ʾ�����Զ����
        "modRemoveManualNo" _        ' ʾ����ɾ���ֹ����
    )
    clsNames = Array()               ' ������ģ�飬��������
    frmNames = Array( _
        "ProgressForm", _            ' ʾ�������ȴ���
        "PageSettings"               ' ʾ����ҳ�洰��
    )

    '�������� ���·�����ļ���
    Dim outDir As String, outName As String, outFull As String, stamp As String
    outDir = Environ$("USERPROFILE") & "\Desktop\������\"    ' Ĭ������/������/
    EnsureFolderExists outDir
    stamp = Format(Now, "yyyymmdd_HHNNSS")
    outName = "Word��ʽ������_" & stamp & ".dotm"
    outFull = outDir & outName

    '����������ѡ���������ʱĿ¼
    Dim tmpRoot As String
    tmpRoot = Environ$("TEMP") & "\vba_pack_" & stamp
    EnsureFolderExists tmpRoot
    Dim dirMod$, dirCls$, dirFrm$
    dirMod = tmpRoot & "\Modules": EnsureFolderExists dirMod
    dirCls = tmpRoot & "\Classes": EnsureFolderExists dirCls
    dirFrm = tmpRoot & "\Forms":   EnsureFolderExists dirFrm

    Application.NormalTemplate.Save
    Dim vbproj As Object: Set vbproj = Application.NormalTemplate.VBProject
    Call ExportComponents(vbproj, modNames, dirMod, 1) ' 1=��׼ģ��
    Call ExportComponents(vbproj, clsNames, dirCls, 2) ' 2=��ģ��
    Call ExportComponents(vbproj, frmNames, dirFrm, 3) ' 3=����

    '���ģ��½��ĵ� �� ������� �� ���Ϊ .dotm
    Dim doc As Document: Set doc = Documents.Add
    Dim target As Object: Set target = doc.VBProject
    Call ImportComponents(target, dirMod, "*.bas")
    Call ImportComponents(target, dirCls, "*.cls")
    Call ImportComponents(target, dirFrm, "*.frm")

    '���壩д�롰����ģ�顱��������ť / �԰�װ / ж�أ�
    Call InjectBootstrapModule(target)

    '����������Ϊ .dotm ģ�岢�ر�
    doc.SaveAs2 FileName:=outFull, FileFormat:=wdFormatXMLTemplateMacroEnabled
    doc.Close SaveChanges:=True

    MsgBox "�����ɣ�" & vbCrLf & outFull, vbInformation
    Exit Sub

ERRH:
    MsgBox "���ʧ�ܣ�" & Err.Number & "���� " & Err.Description, vbCritical
End Sub

'------------------------------------------
' ���������������鵼��ָ���������
' compType: 1=��׼ģ�� 2=��ģ�� 3=����
'------------------------------------------
Private Sub ExportComponents(ByVal vbproj As Object, ByVal names As Variant, _
                             ByVal toDir As String, ByVal compType As Long)
    Dim i As Long, nm$, comp As Object, out$
    If IsEmpty(names) Then Exit Sub
    For i = LBound(names) To IBound(names)
        nm = CStr(names(i))
        On Error Resume Next
        Set comp = vbproj.VBComponents(nm)
        On Error GoTo 0
        If Not comp Is Nothing Then
            If comp.Type = compType Then
                Select Case compType
                    Case 1: out = toDir & "\" & nm & ".bas"
                    Case 2: out = toDir & "\" & nm & ".cls"
                    Case 3: out = toDir & "\" & nm & ".frm" ' .frx ���Զ�һ�𵼳�
                End Select
                comp.Export out
            Else
                Debug.Print "[����] ���Ͳ�ƥ�䣺"; nm
            End If
        Else
            Debug.Print "[δ�ҵ�] �����"; nm
        End If
        Set comp = Nothing
    Next
End Sub

'------------------------------------------
' ���룺���ļ������ .bas/.cls/.frm ���뵽Ŀ�� VBProject
'------------------------------------------
Private Sub ImportComponents(ByVal vbproj As Object, ByVal fromDir As String, ByVal pattern As String)
    Dim f As String
    f = Dir(fromDir & "\" & pattern)
    Do While Len(f) > 0
        vbproj.VBComponents.Import fromDir & "\" & f
        f = Dir
    Loop
End Sub

'------------------------------------------
' д�롰����ģ�顱����ť/�԰�װ/ж��/AutoExec��
' ˵����
' ��һ������ģ��ʱ���Զ�����һ������ʽ�����䡱���������ڡ������ѡ��г��֣�
' ��������ťʾ���󶨵������_ȫ�ָ�ʽ��_������Selection������ģ�
' �������ṩ����װ/ж�ص� Word �����ļ��С���һ�����
'------------------------------------------
'------------------------------------------
Private Sub InjectBootstrapModule(ByVal vbproj As Object)
    Dim c As Object: Set c = vbproj.VBComponents.Add(1) ' 1=��׼ģ��
    c.name = "modBootstrap"
    
    Dim s As String
    
    '��һ������/�˳�����
    s = s & "'��һ������ʱ������ť" & vbCrLf
    s = s & "Public Sub AutoExec()" & vbCrLf
    s = s & "    CreateToolBar" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "'�������˳�ʱ����ť����ѡ��" & vbCrLf
    s = s & "Public Sub AutoExit()" & vbCrLf
    s = s & "    DeleteToolBar" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    '����������/ɾ���������밴ť
    s = s & "'�����������������밴ť����ť����������ܹ��̣�" & vbCrLf
    s = s & "Public Sub CreateToolBar()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.CommandBars(""��ʽ������"").Delete" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    Dim cb As CommandBar" & vbCrLf
    s = s & "    Set cb = Application.CommandBars.Add(Name:=""��ʽ������"", Position:=msoBarTop, Temporary:=True)" & vbCrLf
    s = s & "    Dim btn As CommandBarButton" & vbCrLf
    s = s & "    Set btn = cb.Controls.Add(Type:=msoControlButton)" & vbCrLf
    s = s & "    btn.Caption = ""���һ����ʽ��""" & vbCrLf
    s = s & "    btn.Style = msoButtonIconAndCaption" & vbCrLf
    s = s & "    btn.FaceId = 1085" & vbCrLf
    s = s & "    btn.OnAction = ""���_ȫ�ָ�ʽ��_������Selection""" & vbCrLf
    s = s & "    cb.Visible = True" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Public Sub DeleteToolBar()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Application.CommandBars(""��ʽ������"").Delete" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    '���ģ�һ����װ�������ļ���
    s = s & "'���ģ�һ����װ�����Ʊ�ģ�嵽 Word ����Ŀ¼���������Զ����أ�" & vbCrLf
    s = s & "Public Sub ��װ�������ļ���()" & vbCrLf
    s = s & "    Dim startup As String" & vbCrLf
    s = s & "    startup = Options.DefaultFilePath(wdStartupPath)" & vbCrLf
    s = s & "    If Len(Dir$(startup, vbDirectory)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""δ�ҵ������ļ��У�"" & startup, vbExclamation: Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim src As String, dst As String, nm As String" & vbCrLf
    s = s & "    src = ThisDocument.FullName" & vbCrLf
    s = s & "    nm = CreateObject(""Scripting.FileSystemObject"").GetFileName(src)" & vbCrLf
    s = s & "    dst = startup & IIf(Right$(startup,1)=""\"" ,"""" ,""\"" ) & nm" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    FileCopy src, dst" & vbCrLf
    s = s & "    MsgBox ""�Ѱ�װ����"" & vbCrLf & dst & vbCrLf & ""���� Word ��Ч��"", vbInformation, ""��װ�ɹ�""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    MsgBox ""��װʧ�ܣ�"" & Err.Number & ""����"" & Err.Description, vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    '���壩һ��ж��
    s = s & "'���壩һ��ж�أ��������ļ���ɾ��ͬ��ģ��" & vbCrLf
    s = s & "Public Sub ж�ش������ļ���()" & vbCrLf
    s = s & "    Dim startup As String" & vbCrLf
    s = s & "    startup = Options.DefaultFilePath(wdStartupPath)" & vbCrLf
    s = s & "    Dim nm As String: nm = CreateObject(""Scripting.FileSystemObject"").GetFileName(ThisDocument.FullName)" & vbCrLf
    s = s & "    Dim p As String: p = startup & IIf(Right$(startup,1)=""\"" ,"""" ,""\"" ) & nm" & vbCrLf
    s = s & "    If Len(Dir$(p)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""����Ŀ¼��δ�ҵ���"" & vbCrLf & p, vbExclamation: Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo EH2" & vbCrLf
    s = s & "    Kill p" & vbCrLf
    s = s & "    MsgBox ""��ж�أ�"" & vbCrLf & p & vbCrLf & ""���� Word ��Ч��"", vbInformation" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH2:" & vbCrLf
    s = s & "    MsgBox ""ж��ʧ�ܣ�"" & Err.Number & ""����"" & Err.Description, vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf
    
    c.codeModule.AddFromString s
End Sub

