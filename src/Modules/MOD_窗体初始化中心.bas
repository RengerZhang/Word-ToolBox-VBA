Attribute VB_Name = "MOD_�����ʼ������"
Option Explicit
' =========================================================
'  MOD_��ʽ��ʼ������
'  ���ã�ͳһ���������䴰�塿����ҳ��ĳ�ʼ���������أ�
'  Լ����
'   1��MultiPage �ؼ����Ƽ���mpTabs�������ǣ�Ҳ���Զ�ʶ��
'   2����ҳ Page.Name��
'        - pgPageSetup     ҳ������
'        - pgCaption       ͼ�����
'        - pgTableFormat   ����ʽ��
'        - pgTitle         �������ã�ռλ��
'        - pgStyleImport   ��ʽ���루ռλ��
'   3��ÿҳ�����״λ� force=True ʱִ�У��� Page.Tag="inited" ��ǣ�
' =========================================================


'��һ��������ڣ�����ҳ�����ơ���ʼ���������أ�
Public Sub Init_ByPageName(ByVal host As Object, ByVal pageName As String, Optional ByVal force As Boolean = False)
    Select Case LCase$(pageName)
        Case "pgpagesetup":    Init_PageSetup host, force
        Case "pgcaption":      Init_Caption host, force
        Case "pgtableformat":  Init_TableFormat host, force
        Case "pgtitle":        Init_Title host, force            ' ռλ
        Case "pgstyleimport":  Init_StyleImport host, force      ' ռλ
        Case Else
            ' δ֪ҳ����������
    End Select
End Sub


'������������ڣ���ʼ������ǰѡ�е�ҳ�桱�����ڴ��� Initialize / MultiPage_Change��
Public Sub Init_CurrentPage(ByVal host As Object, Optional ByVal force As Boolean = False)
    Dim mp As Object: Set mp = FindMultiPage(host)
    If mp Is Nothing Then Exit Sub
    Dim curName As String
    curName = mp.Pages(mp.Value).name
    Init_ByPageName host, curName, force
End Sub


'��������ѡ��ڣ�һ���Գ�ʼ��ȫ��ҳ�棨��Ҫ���ָ�Ĭ��ȫ��ҳ�桱ʱ���ã�
Public Sub Init_All(ByVal host As Object, Optional ByVal force As Boolean = False)
    Init_PageSetup host, force
    Init_Caption host, force
    Init_TableFormat host, force
    Init_Title host, force
    Init_StyleImport host, force
End Sub


' =========================================================
'  ����ҳ���ʼ����ÿҳ��һ�Σ����ڵ������ԣ�
' =========================================================

'���ģ�ҳ������ҳ�����д����Ĭ��ֵ�����ٵ��� ps_Init��
Public Sub Init_PageSetup(ByVal host As Object, Optional ByVal force As Boolean = False)
      If Not force Then
        Dim pg As Object: Set pg = GetPage(host, "pgPageSetup")
        If Not pg Is Nothing Then
            If CStr(pg.tag) = "inited" Then Exit Sub
        End If
    End If
    On Error Resume Next     '��һ���ؼ�ȱʧʱ�Զ����������ⱨ��

    ' ���� 1. ����߾ࣨcm������
    host.txtTop.text = "2.5"
    host.txtBottom.text = "2.5"
    host.txtLeft.text = "3"
    host.txtRight.text = "3"

    ' ���� 2. ���߾ࣨcm������
    host.txtTopL.text = "3"
    host.txtBottomL.text = "3"
    host.txtLeftL.text = "2.5"
    host.txtRightL.text = "2.5"

    ' ���� 3. ҳü���� ������������� / �Ҳ�һ�У�
    host.txtHeaderLeft.text = "������������ MHP0-1403��Ԫ" & vbCrLf & "73-04�ؿ�����(��Ǩ)����ס����Ŀ"
    host.txtHeaderRight.text = "ʩ����֯���"

    ' ���� 4. ҳü/ҳ�ž��루cm������
    host.txtHeaderDist.text = "1.5"
    host.txtFooterDist.text = "1.75"

    ' ���� 5. Logo ·����Ĭ�����գ���ǿ�����
    If Len(host.txtLogo.text) = 0 Then host.txtLogo.text = "C:\Users\Tony Zhang\Desktop\logo.png"

    On Error GoTo 0

    MarkInited host, "pgPageSetup"
End Sub

'���壩ͼ�����ҳ����ʼ����ģʽѡ��/��������/�ֺš����������ж���ǿ��ִ�У�
Public Sub Init_Caption(ByVal host As Object, Optional ByVal force As Boolean = False)
    On Error Resume Next   '����ֹ�����ؼ�ȱʧʱ�����жϣ�

    '��һ��ģʽѡ�񣺽������ֹ���䣨DropDownList + MatchRequired��
    With host.cboModeSelect
        SetAsDropDownListSafe host, .name
        .Clear
        .AddItem "��ģʽ"
        .AddItem "ͼģʽ"
        .ListIndex = 0          ' Ĭ�ϣ���ģʽ
        .ListRows = 6
    End With

    '�����������������������ü��ϣ���ֹ����
    With host.cboCapFontCN
        SetAsDropDownListSafe host, .name
        .Clear
        AddChineseFonts .Object ' ������/�����ṩ����亯��
        .Value = "����"
        .ListRows = 12
    End With

    '�������ֺ������������ֺ� + ���ð�ֵ����ֹ����
    With host.cboCapFontSize
       .Style = fmStyleDropDownList
        .MatchRequired = True
        .Clear
        AddChineseFontSizes .Object ' �����е��ֺ���亯��
        .Value = "���"
        .ListRows = 12
    End With

    On Error GoTo 0
End Sub




'����������ʽ��ҳ����ǰ�����Ĭ��ֵ/�����������ֺ� + ���ð�ֵ��
Public Sub Init_TableFormat(ByVal host As Object, Optional ByVal force As Boolean = False)
      If Not force Then
        Dim pg As Object: Set pg = GetPage(host, "pgPageSetup")
        If Not pg Is Nothing Then
            If CStr(pg.tag) = "inited" Then Exit Sub
        End If
    End If
    
    ' ���� 1) ȫ�����ã��ֺ����� ����
    With host.cboFontSize
        SetAsDropDownListSafe host, .name
        .Clear
        AddChineseFontSizes .Object
        .Value = "���"
        .ListRows = 12
    End With

    ' ���� 2) ��ǰ����ֺ����� ����
    With host.cboCurFontSize
        SetAsDropDownListSafe host, .name
        .Clear
        AddChineseFontSizes .Object
        .Value = "���"
        .ListRows = 12
    End With

    ' ���� 3) ȫ�ĸ�ʽ����Ĭ�Ͽ��� ����
    host.chkThickOuter.Value = True        ' ���Ӵ֣���
    host.chkFirstRowBold.Value = True      ' ���мӴ֣���

    ' ���� 4) ��ǰ������ã�Ĭ�Ͽ��� ����
    host.chkCurThickOuter.Value = True     ' ���Ӵ֣���
    host.chkCurFirstRowBold.Value = False  ' ���мӴ֣���
    host.chkCurHeaderRepeat.Value = True   ' �����ظ�����
    host.chkCurAllowBreak.Value = False    ' ���ж�ҳ����

    MarkInited host, "pgTableFormat"
End Sub


'���ߣ���������ҳ��ռλ��������������Ѿ����ʼ���������
Public Sub Init_Title(ByVal host As Object, Optional ByVal force As Boolean = False)
      If Not force Then
        Dim pg As Object: Set pg = GetPage(host, "pgPageSetup")
        If Not pg Is Nothing Then
            If CStr(pg.tag) = "inited" Then Exit Sub
        End If
    End If
    ' TODO������ҳ�ؼ���ʼ��������/Ĭ��ֵ/������
    MarkInited host, "pgTitle"
End Sub


'���ˣ���ʽ����ҳ��ռλ��������������Ѿ����ʼ���������
Public Sub Init_StyleImport(ByVal host As Object, Optional ByVal force As Boolean = False)
    If Not PageNeedsInit(host, "pgStyleImport", force) Then Exit Sub
    ' TODO����ʽ����ҳ��ʼ����Ĭ��·�����б���ť״̬�ȣ�
    MarkInited host, "pgStyleImport"
End Sub


' =========================================================
'  ˽�й��ߣ��뱣������ҳ�����/��ʼ�����/��ȫ����/ͨ������/��ȫ����
' =========================================================

'���ţ����� MultiPage����������Ϊ mpTabs������ȡ��һ�� MultiPage��
Private Function FindMultiPage(ByVal host As Object) As Object
    On Error Resume Next
    Set FindMultiPage = host.Controls("mpTabs")
    If FindMultiPage Is Nothing Then
        Dim ctl As Object
        For Each ctl In host.Controls
            If TypeName(ctl) = "MultiPage" Then
                Set FindMultiPage = ctl
                Exit For
            End If
        Next
    End If
End Function

'��ʮ��ȡ Page�������ڷ��� Nothing��
Private Function GetPage(ByVal host As Object, ByVal pageName As String) As Object
    Dim mp As Object: Set mp = FindMultiPage(host)
    If mp Is Nothing Then Exit Function
    On Error Resume Next
    Set GetPage = mp.Pages(pageName)
End Function

'��ʮһ���Ƿ���Ҫ��ʼ���������ڼ�����ʼ����force=True ǿ�ƣ�
Private Function PageNeedsInit(ByVal host As Object, ByVal pageName As String, ByVal force As Boolean) As Boolean
    Dim pg As Object: Set pg = GetPage(host, pageName)
    If pg Is Nothing Then Exit Function
    If force Then
        PageNeedsInit = True
    Else
        PageNeedsInit = (CStr(pg.tag) <> "inited")
    End If
End Function

'��ʮ�������Ϊ�ѳ�ʼ��
Private Sub MarkInited(ByVal host As Object, ByVal pageName As String)
    Dim pg As Object: Set pg = GetPage(host, pageName)
    If Not pg Is Nothing Then pg.tag = "inited"
End Sub

'��ʮ������ ComboBox ��Ϊ���������� + ����Ϊ�б������ȫ���ã�
Private Sub SetAsDropDownListSafe(ByVal host As Object, ByVal comboName As String)
    On Error Resume Next
    With host.Controls(comboName)
        .Style = fmStyleDropDownList
        .MatchRequired = True
    End With
End Sub

'��ʮ�ģ�ͨ�ã���䡰�����ֺ� + ���ð�ֵ������ʾ�ת��ֵ�����㹫������ GetFontSizePt��
Private Sub AddChineseFontSizes(ByVal cbo As Object)
    With cbo
        .AddItem "����": .AddItem "С��": .AddItem "һ��": .AddItem "Сһ"
        .AddItem "����": .AddItem "С��": .AddItem "����": .AddItem "С��"
        .AddItem "�ĺ�": .AddItem "С��": .AddItem "���": .AddItem "С��"
        .AddItem "����": .AddItem "С��"
        .AddItem "8": .AddItem "9": .AddItem "10": .AddItem "12"
        .AddItem "14": .AddItem "16": .AddItem "18"
    End With
End Sub
'��ʮ��-B��ͨ�ã���䳣������/�������壨�ɰ�����չ��
Private Sub AddChineseFonts(ByVal cbo As Object)
    With cbo
        .AddItem "����"
        .AddItem "����"
        .AddItem "����"
        .AddItem "����"
        .AddItem "΢���ź�"
        .AddItem "Times New Roman"  ' ��ĸ/���ֳ���
    End With
End Sub


'��ʮ�壩��ȫ���ã����巽����������ã��� CapPage_Init��
Private Function TryCallHostMethod(ByVal host As Object, ByVal methodName As String, ByVal force As Boolean) As Boolean
    On Error Resume Next
    CallByName host, methodName, VbMethod, force
    TryCallHostMethod = (Err.Number = 0)
    Err.Clear
End Function

'��ʮ������ȫ���ã�ģ����̴��������У�֧�֡�ģ��.�����������������������
Private Function RunIfExists(procFullName As String, ByVal host As Object, ByVal force As Boolean) As Boolean
    On Error Resume Next
    Application.Run procFullName, host, force
    RunIfExists = (Err.Number = 0)
    Err.Clear
End Function

