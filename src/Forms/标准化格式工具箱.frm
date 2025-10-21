VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��׼����ʽ������ 
   Caption         =   "��׼����ʽ������  V1.0.250919"
   ClientHeight    =   5270
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9170.001
   OleObjectBlob   =   "��׼����ʽ������.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��׼����ʽ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ���� ���巵��ֵ�������� Show ���ȡ������
Public SelectedThickOuter As Boolean      ' ���1.5�����أ�ȫ�ģ�
Public SelectedFirstRowBold As Boolean    ' ���мӴֿ��أ�ȫ�ģ�
Public SelectedFontSizeName As String     ' �����ֺ�����ȫ�ģ�
Public SelectedFontSizePt As Single       ' ��Ӧ��ֵ��ȫ�ģ�
Public Canceled As Boolean                ' �Ƿ�ȡ��


' =========================
'  �����ʼ����ȫ����ʼ����
' =========================
Private Sub UserForm_Initialize()
    '��һ��ȫ����ҳ��һ���Գ�ʼ����ǿ������Ĭ�ϣ�
    Init_All Me, True

    '��������Ҫ�󡿱���ʽ��ҳ�Ŀؼ�Ĭ��ֵ���ƶ�������ʼ�����ġ���
    '     ��ԭ��������� chk/cbo ��Ĭ�ϸ�ֵ����ɾ���������ظ���
End Sub

' =========================
'  ͨ��С����
' =========================
'��һ���ж��Ƿ��ڡ������� + MultiPage����
Private Function InToolbox() As Boolean
    On Error Resume Next
    Dim c As MSForms.Control
    Set c = Me.Controls("mpTabs")    ' �Ƽ� MultiPage ��
    InToolbox = (Err.Number = 0)
    Err.Clear
End Function

'��������ȫ���ã����巽����������ã��� CapPage_Init��
Private Function TryCallHostMethod(ByVal host As Object, ByVal methodName As String, ParamArray args()) As Boolean
    On Error Resume Next
    CallByName host, methodName, VbMethod, args
    TryCallHostMethod = (Err.Number = 0)
    Err.Clear
End Function

'��������ȫ���ã�ģ����̴��������У�֧�֡�ģ��.�����������������������
Private Function RunIfExists(procFullName As String, ParamArray args()) As Boolean
    On Error Resume Next
    Application.Run procFullName, args
    RunIfExists = (Err.Number = 0)
    Err.Clear
End Function

'���ģ��ѡ�������ҳ�����á����ֵͬ�����ɴ��壨���ƣ���ʽ_��׼��ҳ�����ã�
Private Sub CopyPageSetupToOldForm(ByVal oldForm As Object)
    On Error Resume Next
    oldForm.txtTop.text = Me.txtTop.text
    oldForm.txtBottom.text = Me.txtBottom.text
    oldForm.txtLeft.text = Me.txtLeft.text
    oldForm.txtRight.text = Me.txtRight.text

    oldForm.txtTopL.text = Me.txtTopL.text
    oldForm.txtBottomL.text = Me.txtBottomL.text
    oldForm.txtLeftL.text = Me.txtLeftL.text
    oldForm.txtRightL.text = Me.txtRightL.text

    oldForm.txtHeaderLeft.text = Me.txtHeaderLeft.text
    oldForm.txtHeaderRight.text = Me.txtHeaderRight.text
    oldForm.txtLogo.text = Me.txtLogo.text
    oldForm.txtHeaderDist.text = Me.txtHeaderDist.text
    oldForm.txtFooterDist.text = Me.txtFooterDist.text
    On Error GoTo 0
End Sub


' =========================================================
'  һ������ʽ���롿ҳ��ռλ����������ʵʵ�ֲ�������
' =========================================================

Private Sub cmdStyleImport_Click()
    Call һ������ȫ����ʽ
End Sub


' =========================================================
'  �������������á�ҳ��ռλ����������ʵ�߼���
' =========================================================
Private Sub cmdAutoDetectHeading_Click()
    Call ƥ����Ⲣ������ʽ_������������
End Sub
Private Sub cmdMultiLevelMatch_Click()
    Call �����Զ����
End Sub
Private Sub cmdRemoveManualNumber_Click()
'    Call ȥ���ֹ����_������������
    Call ȥ���ֹ����_ʹ�ý��ȴ���
End Sub


' =========================
' ������ҳ������ҳ������ɴ��塰�н���׼��ҳ�����á�����
' =========================
' ���� ҳ������ �� ��� LOGO ����
Private Sub cmdBrowse_Click()
    '��һ��ֱ���þɴ���ִ������߼����������������ͬ����ȥ
    ��ʽ_��׼��ҳ������.PS_BrowseLogo Me
End Sub

' ���� ҳ������ �� Ӧ�õ����� ����
Private Sub cmdApplySection_Click()
    '������ͬ������ �� �ɴ���ִ��ҵ��
    ��ʽ_��׼��ҳ������.PS_ApplySection Me
End Sub

' ���� ҳ������ �� ȫ��Ӧ�� ����
Private Sub cmdApplyAll_Click()
    '������ͬ������ �� �ɴ���ִ��ҵ�񣨺����ȣ�
    ��ʽ_��׼��ҳ������.PS_ApplyAll Me
End Sub


' =========================================================
'  �ġ�������ʽ����ҳ
' =========================================================
' ���� ȫ������OK = ֱ���ܡ�ȫ�ı���ʽ�������������� ����
Private Sub cmdOK_Click()
    '��һ��ȡ�����������
    Dim nm As String, pt As Single
    nm = Trim(Me.cboFontSize.text)
    If Len(nm) = 0 Then
        MsgBox "��ѡ�������ֺţ��硰��š������������ְ�ֵ��", vbExclamation
        Exit Sub
    End If

    pt = GetFontSizePt(nm)   ' ��Ĺ�������
    If pt <= 0 Then
        MsgBox "�ֺ���Ч��" & nm, vbExclamation
        Exit Sub
    End If

    '������ֱ��ִ�У����ٵ��� dlg�����رձ����壩
    ȫ�ı���ʽ��_������ _
        Me.chkThickOuter.Value, _
        Me.chkFirstRowBold.Value, _
        pt, _
        nm

    '������������ʾ�ɼӣ�
    ' MsgBox "ȫ�ı���ʽ������ɡ�", vbInformation
End Sub
Private Sub CommandButton23_Click()
    '��һ��ȡ�����������
    Dim nm As String, pt As Single
    nm = Trim(Me.cboFontSize.text)
    If Len(nm) = 0 Then
        MsgBox "��ѡ�������ֺţ��硰��š������������ְ�ֵ��", vbExclamation
        Exit Sub
    End If

    pt = GetFontSizePt(nm)   ' ��Ĺ�������
    If pt <= 0 Then
        MsgBox "�ֺ���Ч��" & nm, vbExclamation
        Exit Sub
    End If

    '������ֱ��ִ�У����ٵ��� dlg�����رձ����壩
    ȫ�ı���ʽ��_������1 _
        Me.chkThickOuter.Value, _
        Me.chkFirstRowBold.Value, _
        pt, _
        nm

    '������������ʾ�ɼӣ�
    ' MsgBox "ȫ�ı���ʽ������ɡ�", vbInformation
End Sub

Private Sub cmdTF_ApplyAll_Click()
    '������ȫ�ı���ʽ����ռλ����������ܿع��̣�
    ' �߼�˳��
    '  1) ��ȡ��ȫ�����򡱵�����ѡ����Ӵ�/���мӴ�/�ֺ�
    '  2) ��������ܿع��̣����磺ȫ�ı���ʽ�����ߣ�����д��ȫ���������ܿض�ȡ
    '  3) ���ȴ���/�쳣����
    Dim nm As String, pt As Single
    nm = Trim(cboFontSize.text)
    pt = GetFontSizePt(nm)
    If pt <= 0 Then
        MsgBox "����ѡ����Ч�������ֺŻ����ְ�ֵ��", vbExclamation: Exit Sub
    End If

    ' ��ռλ���á������еĴ���̣���������ռλ��ʾ
    If Not RunIfExists("ȫ�ı���ʽ������") Then
        MsgBox "��ռλ���뽫ȫ�ı���ʽ��������������Ϊ��ȫ�ı���ʽ�����ߡ������ڴ˴���Ϊ��Ĺ�������", vbInformation
    End If
End Sub

' ���� ��ǰ������ã�������ʵ�֣��������������� ����
Private Sub cmdApplyCur_Click()
    '��������ǰ������� �� ���������й��̡���ǰ����ʽ���á�
    Dim pt As Single
    pt = GetFontSizePt(Me.cboCurFontSize.text)
    If pt <= 0 Then
        MsgBox "����ѡ����Ч�������ֺŻ����ְ�ֵ��", vbExclamation
        Exit Sub
    End If

    Call ��ǰ����ʽ����( _
        Me.chkCurThickOuter.Value, _
        Me.chkCurFirstRowBold.Value, _
        Me.chkCurHeaderRepeat.Value, _
        Me.chkCurAllowBreak.Value, _
        pt)
End Sub

Private Sub cmdTF_Explain_Click()
    '���ģ�����˵����ռλ��
    MsgBox "��ռλ��չʾ��ȫ�ı���ʽ������˵����ע�����", vbInformation
End Sub


' =========================================================
'  �塢��ͼ����⡿ҳ���¼�ռλ��

'==============================
' ��һ����ť�¼� �� ͳһ����
'==============================
Private Sub btnAllMatchStyles_Click()
    ����ִ�� 1   '��1��������ʽƥ��
End Sub
Private Sub btnAllAutoNumber_Click()
    ����ִ�� 2   '��2���Զ�ͼ����
End Sub
Private Sub btnAllRemoveManualNo_Click()
    ����ִ�� 3   '��3��ȥ���ֹ����
End Sub
Private Sub btnCaptionPreCheck_Click()
    ����ִ�� 4   '��4������Ԥ��飺��/ͼ����ģʽ����
End Sub
Private Sub btnCheckPictures_Click()
    ����ִ�� 5
End Sub
'==============================
' ���������ĵ��ȣ�����ģʽ������ A/B/C �� D/E/F
'==============================
Private Sub ����ִ��(ByVal ������� As Long)
    Dim modeKey As String
    modeKey = ��ȡģʽKey(Me)   ' ���� "��" | "ͼ" | ""
    If modeKey = "" Then
        MsgBox "δѡ����Чģʽ�����ڡ�ģʽѡ��������ѡ�񡾱�ģʽ/ͼģʽ����", vbExclamation
        Exit Sub
    End If

    Select Case modeKey
        Case "��"
            Select Case �������
                Case 1: Call �������ʽͳһ
                Case 2: Call �������Զ����_ʹ�ý��ȴ���
                Case 3: Call ��������ֹ����_ʹ�ý��ȴ���
                Case 4: Call �Լ�_��������ʽһ����
                Case Else: GoTo MAP_ERR
            End Select

        Case "ͼ"
            Select Case �������
                Case 1: Call ͼƬ������ʽͳһ_������
                Case 2: Call ͼƬ�����Զ����_ʹ�ý��ȴ���
                Case 3: Call ���ͼƬ���ֹ����_ʹ�ý��ȴ���
                Case 4: Call �Լ�_ͼƬ������ʽһ����
                Case 5: Call ͳһͼƬ������ʽ_ʹ�ý��ȴ���
                Case Else: GoTo MAP_ERR
            End Select
    End Select
    Exit Sub

MAP_ERR:
    MsgBox "δӳ�䵽��Ӧ�Ķ��������鰴ť�����ģʽӳ�䡣", vbCritical
End Sub

'===============================
' ���룺���������������ǵ�Ŀ����ʽ�����½���
' ��ģʽ �� �����⣻ͼģʽ �� ͼƬ����
'===============================
Private Sub btnCapImport_Click()
    Dim doc As Document: Set doc = ActiveDocument
    Dim modeKey As String, styleName As String
    Dim sty As Style
    Dim fontCN As String, sizeText As String
    Dim sizePt As Single, beforeLines As Single, afterLines As Single
    Dim oneLinePt As Single, boldOn As Boolean
    
    '��һ���ж�ģʽ��ֻ�����֡���/ͼ����
    modeKey = ��ȡģʽKey(Me)                 ' ��ǰ���Ѽӹ��Ĺ��ߺ���
    If modeKey = "" Then
        MsgBox "�����ڡ�ģʽѡ����ѡ�� ��ģʽ/ͼģʽ��", vbExclamation
        Exit Sub
    End If
    styleName = IIf(modeKey = "��", "������", "ͼƬ����")
    
    '��������ȡĿ����ʽ�����½���
    On Error Resume Next
    Set sty = doc.Styles(styleName)
    On Error GoTo 0
    If sty Is Nothing Then
        MsgBox "��" & styleName & "�������ڣ������ڡ���ʽ���롿�����е����һ�����롿��", vbExclamation
        Exit Sub
    End If
    
    '��������ȡ������
    fontCN = NzStr(Me.cboCapFontCN.Value, "����")
    sizeText = NzStr(Me.cboCapFontSize.Value, "���")
    sizePt = �ֺŵ���ֵ(sizeText, 10.5!)        ' Ĭ����š�10.5pt
    boldOn = (Me.chkCapBold.Value = True)
    beforeLines = val(��׼�������ı�(Me.txtParaSpaceBeforeLines.text)) ' �����=0
    afterLines = val(��׼�������ı�(Me.txtParaSpaceAfterLines.text))
    
    '���ģ��ԡ��ֺš���Ϊ 1 �еĻ�׼ pt ֵ����������������ġ��С����
    oneLinePt = sizePt
    
    '���壩���Ǳ�Ҫ���ԣ����಻����
    With sty.Font
        .NameFarEast = fontCN
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .bold = boldOn
        .Size = sizePt
    End With
    With sty.ParagraphFormat
        .SpaceBefore = beforeLines * oneLinePt
        .SpaceAfter = afterLines * oneLinePt
        ' ���Ķ������뷽ʽ���оࡢ��ټ����������Ʊ�λ��
    End With
    
    '����������ʹ�ø���ʽ�Ķ�������ˢ��һ�Σ����ı��ı���
    ǿ��������ʽ_�� doc, sty
    
    MsgBox "�Ѹ�����ʽ��" & styleName & "����" & vbCrLf & _
           "��������=" & fontCN & "���ֺ�=" & sizeText & "��" & Format(sizePt, "0.0#") & "pt��" & vbCrLf & _
           "�Ӵ�=" & IIf(boldOn, "��", "��") & "����ǰ=" & beforeLines & "�У��κ�=" & afterLines & "�С�", _
           vbInformation
End Sub

' =========================================================
'  ������ͼƬ������ҳ��
' =========================================================
Private Sub ȫ�ı��ʽ��_����ͼƬ_Click()
Call Ԥ����_���ͼƬ������ͨ��
End Sub

Private Sub ����ΪͼƬ��_Click()
Call һ������ѡ�����ΪͼƬ��
End Sub
Private Sub btnCoverage_Click()
Call ���ɷ���_��������
End Sub
Private Sub CommandButton19_Click()
Call ���뵥��ͼƬ��_ͼƬ�ؼ���
End Sub
Private Sub CommandButton22_Click()
Call ����˫��ͼƬ��_ͼƬ�ؼ���_˫��
End Sub


' =========================================================
'  �ߡ����弶����˳�/ȡ��
' =========================================================
Private Sub cmdCancel_Click()
    Canceled = True
    Unload Me
End Sub


'==============================
' ���������ߣ���ȡģʽ�������к�
'==============================
' ��ȡ����ֵ���������� "��"/"ͼ"�����෵�ؿգ�
Private Function ��ȡģʽKey(ByVal host As Object) As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(host.cboModeSelect.Value))
    On Error GoTo 0
    If Len(v) = 0 Then Exit Function
    v = Replace$(v, "ģʽ", "")         ' ȥ����ģʽ�����֣��磺��ģʽ/ͼģʽ��
    v = Left$(v, 1)                     ' ֻȡ����
    If v = "��" Or v = "ͼ" Then ��ȡģʽKey = v
End Function

' ���������� Public Sub���ڱ�׼ģ���У����������Ѻô�����ʾ
Private Sub ��ȫ���к�(ByVal macroName As String)
    On Error GoTo EH
    Application.Run macroName
    Exit Sub
EH:
    MsgBox "δ�ҵ���ִ�й��̣�" & macroName & vbCrLf & _
           "��ȷ�ϸù��̴����ڡ���׼ģ�顱��Ϊ Public Sub��", vbExclamation
End Sub
'���������ߣ���/Null ���Ĭ���ַ���
Private Function NzStr(ByVal v As Variant, ByVal def As String) As String
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then NzStr = def Else NzStr = CStr(v)
End Function

'���������ߣ��ѡ�10.5 / �������� / 10.5pt�������ı���׼��Ϊ�� Val �İ������
Private Function ��׼�������ı�(ByVal s As String) As String
    s = Trim$(s)
    s = Replace$(s, "��", "."): s = Replace$(s, "��", ".")
    s = Replace$(s, "��", "."): s = Replace$(s, "��", ".")
    s = Replace$(s, "��", "-"): s = Replace$(s, "��", "-")
    s = Replace$(s, "��", "+")
    s = Replace$(s, "pt", "", , , vbTextCompare)
    s = Replace$(s, "�У�", "", , , vbTextCompare)
    ��׼�������ı� = s
End Function

'���������ߣ������ֺ� �� pt����δʶ��������ֵ�����ջ���Ĭ��ֵ
Private Function �ֺŵ���ֵ(ByVal s As String, ByVal defPt As Single) As Single
    Dim key As String: key = Trim$(s)
    Select Case key
        Case "����": �ֺŵ���ֵ = 42#
        Case "С��": �ֺŵ���ֵ = 36#
        Case "һ��": �ֺŵ���ֵ = 26#
        Case "Сһ": �ֺŵ���ֵ = 24#
        Case "����": �ֺŵ���ֵ = 22#
        Case "С��": �ֺŵ���ֵ = 18#
        Case "����": �ֺŵ���ֵ = 16#
        Case "С��": �ֺŵ���ֵ = 15#
        Case "�ĺ�": �ֺŵ���ֵ = 14#
        Case "С��": �ֺŵ���ֵ = 12#
        Case "���": �ֺŵ���ֵ = 10.5
        Case "С��": �ֺŵ���ֵ = 9#
        Case "����": �ֺŵ���ֵ = 7.5
        Case "С��": �ֺŵ���ֵ = 6.5
        Case "�ߺ�": �ֺŵ���ֵ = 5.5
        Case "�˺�": �ֺŵ���ֵ = 5#
        Case Else
            ' ����ֱ������ 10.5 / 11 ��
            key = ��׼�������ı�(key)
            If IsNumeric(key) Then
                �ֺŵ���ֵ = CSng(key)
            Else
                �ֺŵ���ֵ = defPt
            End If
    End Select
End Function

'���������ߣ����ĵ�����Ӧ��ĳ��ʽ�Ķ��䡰����һ�Ρ���������Ч������ʽѭ����
Private Sub ǿ��������ʽ_��(ByVal doc As Document, ByVal Ŀ����ʽ As Style)
    With doc.content.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .text = ""
        .replacement.text = ""
        .Style = Ŀ����ʽ
        .replacement.Style = Ŀ����ʽ
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

