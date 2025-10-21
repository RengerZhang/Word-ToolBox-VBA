VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��ʽ_������ʽһ������ 
   Caption         =   "�н���׼����ʽ���ù���"
   ClientHeight    =   9920.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   21030
   OleObjectBlob   =   "��ʽ_������ʽһ������.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��ʽ_������ʽһ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' ǿ�Ʊ�������

Sub ShowMyUserForm()
    UserForm.Show  ' ģ̬��ʾ����
End Sub
Option Explicit
Private Sub btnApply_Click()
    Dim i As Integer
    Dim styleName As String        ' ��ʽ����
    Dim outlineLevel As Integer    ' ��ټ���ComboBox ������
    Dim outlineLevelVal As Integer ' ʵ�ʴ�ټ���0~9��
    Dim fontSizeText As String     ' �ֺ���ʾ�ı�����"С��"��"�ĺ�"��
    Dim fontSizePt As Single       ' ӳ�����ֺŰ�ֵ
    Dim isBold As Boolean          ' �Ƿ�Ӵ�
    Dim alignment As Integer       ' ���뷽ʽ��ö��ֵ��
    Dim fontName As String         ' ��������
    Dim fontAsciiName As String    ' ��������
    Dim indentType As String       ' �������ͣ�����/���ң�
    Dim indentValue As Single      ' ����ֵ���ַ���
    
    ' ���� 1~10 ����ʽ���ã�����ʵ���������ѭ����Χ��
    For i = 1 To 10
        ' ---------------------- 1. ��ȡ�ؼ�ֵ���ؼ����ϸ�ƥ��ؼ������� ----------------------
        ' ��ʽ���ƣ�TextBox01 + ��λ�кţ��� TextBox0101 ~ TextBox0110��
        styleName = Trim(Me.Controls("TextBox01" & Format(i, "00")).Value)
        If styleName = "" Then
            MsgBox "�� " & i & " ����ʽ����Ϊ�գ�����д�����ԣ�", vbExclamation, "�������"
            Exit Sub
        End If
        
        ' ��ټ���ComboBox02 + ��λ�кţ��� ComboBox0201 ~ ComboBox0210��
        If Me.Controls("ComboBox02" & Format(i, "00")).ListIndex < 0 Then
            MsgBox "�� " & i & " �д�ټ���δѡ�������ú����ԣ�", vbExclamation, "�������"
            Exit Sub
        End If
        outlineLevel = Me.Controls("ComboBox02" & Format(i, "00")).ListIndex
        outlineLevelVal = outlineLevel + 1  ' ListIndex �� 0 ��ʼ��ʵ�ʼ��� +1
        
        ' ---------------------- 2. �ֺ�ӳ�䴦������ʵ�֣� ----------------------
        ' ��ȡ�ֺ���ʾ�ı�����"С��"��"12"��
        fontSizeText = Me.Controls("ComboBox03" & Format(i, "00")).Value
        If fontSizeText = "" Then
            MsgBox "�� " & i & " ���ֺ�δѡ�������ú����ԣ�", vbExclamation
            Exit Sub
        End If
        
        ' ����ӳ�亯��ת��Ϊ��ֵ
        fontSizePt = GetFontSizePt(fontSizeText)
        If fontSizePt <= 0 Then
            MsgBox "�� " & i & " ���ֺ���Ч��" & fontSizeText, vbExclamation
            Exit Sub
        End If
        
        
        ' �Ƿ�Ӵ֣�CheckBox04 + ��λ�кţ��� CheckBox0401 ~ CheckBox0410��
        isBold = (Me.Controls("CheckBox04" & Format(i, "00")).Value = True)
        
        ' ���뷽ʽ��ComboBox05 + ��λ�кţ��� ComboBox0501 ~ ComboBox0510��
        alignment = Me.Controls("ComboBox05" & Format(i, "00")).ListIndex
        Select Case alignment
            Case 0: alignment = wdAlignParagraphLeft    ' �����
            Case 1: alignment = wdAlignParagraphCenter  ' ���ж���
            Case 2: alignment = wdAlignParagraphRight   ' �Ҷ���
            Case 3: alignment = wdAlignParagraphJustify ' ���˶���
            Case Else: alignment = wdAlignParagraphLeft ' Ĭ�������
        End Select
        
        ' �������壨ComboBox06 + ��λ�кţ��� ComboBox0601 ~ ComboBox0610��
        fontName = Me.Controls("ComboBox06" & Format(i, "00")).Value
        If fontName = "" Then
            MsgBox "�� " & i & " ����������δѡ�������ú����ԣ�", vbExclamation, "�������"
            Exit Sub
        End If
        
        ' �������壨ComboBox07 + ��λ�кţ��� ComboBox0701 ~ ComboBox0710��
        fontAsciiName = Me.Controls("ComboBox07" & Format(i, "00")).Value
        If fontAsciiName = "" Then
            MsgBox "�� " & i & " ����������δѡ�������ú����ԣ�", vbExclamation, "�������"
            Exit Sub
        End If
        
        ' �������ͣ�ComboBox08 + ��λ�кţ��� ComboBox0801 ~ ComboBox0810��
        indentType = Me.Controls("ComboBox08" & Format(i, "00")).Value
        Select Case indentType
            Case "��������": indentValue = 2  ' ʾ������������ 2 �ַ�
            Case "��������": indentValue = -2 ' ʾ������������ 2 �ַ�����ֵ��ʾ���ң�
            Case Else: indentValue = 0        ' ������
        End Select
        
        ' ---------------------- 2. ����/�޸���ʽ�������߼��� ----------------------
        Dim myStyle As Style
        On Error Resume Next
        Set myStyle = ActiveDocument.Styles(styleName)  ' ���Ի�ȡ������ʽ
        On Error GoTo 0
        
        ' ��ʽ���������½�����������
        If myStyle Is Nothing Then
            Set myStyle = ActiveDocument.Styles.Add( _
                name:=styleName, _
                Type:=wdStyleTypeParagraph _
            )
        End If
        
        ' ---------------------- 3. ������ʽ���ԣ��������ã� ----------------------
        With myStyle
            ' ��ټ���
            .ParagraphFormat.outlineLevel = outlineLevelVal
            
            ' ���壨��/���ģ�
            .Font.name = fontName          ' ��������
            .Font.NameAscii = fontAsciiName ' ��������
            .Font.Size = fontSizePt        ' �ֺţ�����
            .Font.bold = isBold            ' �Ƿ�Ӵ�
            
            ' �������
            .ParagraphFormat.alignment = alignment
            
            ' ����������/���ң�
            .ParagraphFormat.FirstLineIndent = indentValue
            
            ' ����չ���������ԣ����оࡢ��ǰ�κ���ȣ�
            '.ParagraphFormat.LineSpacing = 1.5  ' ʾ����1.5 ���о�
        End With
        
        ' ---------------------- 4. ��ʾ��������ѡ����֪�û����ȣ� ----------------------
        MsgBox "��ʽ '" & styleName & "' ����/�޸ĳɹ���" & vbCrLf & _
               "�� ��ټ���" & outlineLevelVal & vbCrLf & _
               "�� �ֺţ�" & fontSizePt & " ��", vbInformation, "�������"
    Next i
    
    MsgBox "������ʽ��������ɣ�", vbInformation, "�����������"
End Sub


' �û������ʼ����Ϊ�ؼ��ṩ��ʼֵ
Private Sub UserForm_Initialize()
    Dim i As Integer
    For i = 1 To 10
        ' ��ʼ����ټ���������
        With Me.Controls("ComboBox02" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "��"  ' �޴�ټ���ѡ��
            .AddItem "1��"
            .AddItem "2��"
            .AddItem "3��"
            .AddItem "4��"
            .AddItem "5��"
        End With

        ' ��ʼ���ֺ�������
        With Me.Controls("ComboBox03" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "����": .AddItem "С��": .AddItem "һ��": .AddItem "Сһ"
            .AddItem "����": .AddItem "С��": .AddItem "����": .AddItem "С��"
            .AddItem "�ĺ�": .AddItem "С��": .AddItem "���": .AddItem "С��"
            .AddItem "8": .AddItem "9": .AddItem "10": .AddItem "12"
            .AddItem "14": .AddItem "16": .AddItem "18"
        End With
        
        With Me.Controls("ComboBox05" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "�����"
            .AddItem "���ж���"
            .AddItem "�Ҷ���"
            .AddItem "��ɢ����"
        End With

        ' ��ʼ����������������
        With Me.Controls("ComboBox06" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "����"
            .AddItem "����"
            .AddItem "����"
            .AddItem "΢���ź�"
            .AddItem "����"
        End With

        ' ��ʼ����������������
        With Me.Controls("ComboBox07" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "Times New Roman"
            .AddItem "Arial"
            .AddItem "Verdana"
        End With

        ' ��ʼ����������������
        With Me.Controls("ComboBox08" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "��"
            .AddItem "��������"
            .AddItem "��������"
        End With
        
        ' ��ʼ����ǰ�κ��൥λ
        With Me.Controls("ComboBox10" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "��"
            .AddItem "��"
            .AddItem "Ӣ��"
            .AddItem "����"
            .AddItem "����"
        End With
        
        With Me.Controls("ComboBox12" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "��"
            .AddItem "��"
            .AddItem "Ӣ��"
            .AddItem "����"
            .AddItem "����"
        End With
        
        With Me.Controls("ComboBox13" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "�����о�"
            .AddItem "�౶�о�"
            .AddItem "�̶�ֵ"
            .AddItem "��Сֵ"
        End With

        With Me.Controls("ComboBox15" & Format(i, "00"))
            .Clear ' �������ѡ��
            .AddItem "��"
            .AddItem "��"
            .AddItem "Ӣ��"
            .AddItem "����"
            .AddItem "����"
        End With

    Next i
    
End Sub

' �رհ�ť���رմ���
Private Sub btnClose_Click()
    Unload Me
End Sub

' ---------- ���ߺ������ֺ��ı�ת��ֵ ----------
Private Function GetFontSizePt(sizeText As String) As Single
    Dim sizeMap As Object
    Set sizeMap = CreateObject("Scripting.Dictionary")
    With sizeMap
        .Add "����", 42
        .Add "С��", 36
        .Add "һ��", 26
        .Add "Сһ", 24
        .Add "����", 22
        .Add "С��", 18
        .Add "����", 16
        .Add "С��", 15
        .Add "�ĺ�", 14
        .Add "С��", 12
        .Add "���", 10.5
        .Add "С��", 9
        .Add "����", 7.5
        .Add "С��", 6.5
    End With

    If sizeMap.exists(sizeText) Then
        GetFontSizePt = sizeMap(sizeText)
    ElseIf IsNumeric(sizeText) Then
        GetFontSizePt = CSng(sizeText)
    Else
        GetFontSizePt = -1
    End If
End Function

' ---------- ���ߺ�������ȡ��������ֵ ----------
Private Function GetIndentValue(indentType As String) As Single
    If indentType = "��������" Then
        GetIndentValue = 1.5 ' ����������ֵ
    Else
        GetIndentValue = 0 ' ������
    End If
End Function


