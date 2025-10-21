VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "����ʽ���������"
   ClientHeight    =   3500
   ClientLeft      =   90
   ClientTop       =   410
   ClientWidth     =   4990
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProgressForm �Ĵ���
Public Event OnClose()  ' ����ر��¼������ⲿ���ģ�
Public stopFlag As Boolean ' ���ڿ���ǿ��ֹͣ
'
'
'Private Sub TextBoxStatus_Change()
'
'End Sub

Private Sub TextBoxStatus_Change()

End Sub

Private Sub UserForm_Initialize()
    ' ��ʼ���ؼ�
    Me.LabelProgressBar.width = 200
    Me.LabelProgressBar.Height = 24
    Me.LabelProgressBar.Top = 5
    Me.LabelProgressBar.Left = 5
    Me.FrameProgress.width = 0 ' ��ʼ����Ϊ 0
    Me.FrameProgress.Top = 5.5
    Me.FrameProgress.Left = 5.5
    Me.FrameProgress.Height = 25
    Me.TextBoxStatus.width = 235
    Me.TextBoxStatus.Height = 100
    Me.TextBoxStatus.MultiLine = True ' ֧�ֶ���
    Me.TextBoxStatus.ScrollBars = fmScrollBarsVertical ' ���ô�ֱ������
    Me.TextBoxStatus.Locked = True ' ʹ TextBox Ϊֻ��ģʽ
    
    ' ��һ����ʾ�ܱ����������ʼ��ʽ����ʾ
'    Me.TextBoxStatus.Text = "���ĵ��ܹ��� " & ActiveDocument.Tables.Count & " ��������ڿ�ʼ��ʽ��..."
    
    ' ��ʼ��������
    Me.FrameProgress.width = 0
    Me.LabelPercentage.caption = "0%" ' ��ʼ���ٷֱ���ʾ
End Sub
' ����ҳ�湤�ߡ����õĳ�ʼ�����
Public Sub InitForPageSetting(ByVal totalSections As Long, ByVal startSectionIndex As Long)
    Me.caption = "ȫ��ҳ�����ô������"
    Me.FrameProgress.width = 0
    Me.LabelPercentage.caption = "0%"
    Me.TextBoxStatus.text = _
        "���Ĺ��� " & totalSections & " �ڣ����Ĵӵ� " & startSectionIndex & " �ڿ�ʼ��"
End Sub

' ���½�������״̬�ı���
Public Sub UpdateProgressBar(ByVal progress As Integer, ByVal statusMessage As String)
    If stopFlag Then
'        Me.TextBoxStatus.Text = "������ǿ��ֹͣ��"
        progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������ǿ��ֹͣ��"
        Exit Sub
    End If
    
    ' ���½�����
    Me.FrameProgress.width = progress
    ' ���°ٷֱ�
    Me.LabelPercentage.caption = progress / 2 & "%" ' 200px Ϊ�����ȣ������������

    ' ׷�Ӽ�¼�� TextBox
    Me.TextBoxStatus.text = Me.TextBoxStatus.text & vbCrLf & statusMessage
    
   ' ==== ǿ�ƹ��������һ�У��������߰�ť���㣩====
    Dim ctrl As MSForms.Control
    Set ctrl = Me.ActiveControl
    
    Me.TextBoxStatus.SetFocus
    Me.TextBoxStatus.SelStart = Len(Me.TextBoxStatus.text)
    Me.TextBoxStatus.SelLength = 0
    
    ' ��ԭԭ������
    If Not ctrl Is Nothing Then
        On Error Resume Next
        ctrl.SetFocus
        On Error GoTo 0
    End If
    ' ================================================

    
    DoEvents ' ȷ������
End Sub

' ǿ��ֹͣ��ť����¼�
Private Sub btnStop_Click()
    stopFlag = True
'    Me.TextBoxStatus.Text = "�����ѱ�ǿ��ֹͣ��"
    Me.TextBoxStatus.text = Me.TextBoxStatus.text & vbCrLf & "�����ѱ�ǿ��ֹͣ��"
    Me.LabelPercentage.caption = "��ֹ��"
    RaiseEvent OnClose    ' �����ر��¼�
End Sub

' �رհ�ť����¼�
Private Sub btnClose_Click()
    Me.Hide ' ���ش���
    RaiseEvent OnClose    ' �����ر��¼�
End Sub
Private Sub ScrollTextBoxToEnd(tb As MSForms.TextBox)
    Dim ctrl As Control
    ' �ȼ�ס��ǰ����ؼ�
    Set ctrl = Me.ActiveControl
    
    tb.SetFocus
    tb.SelStart = Len(tb.text)
    tb.SelLength = 0
    
    ' �ٰѽ��㻹ԭ��ȥ
    If Not ctrl Is Nothing Then ctrl.SetFocus
End Sub
