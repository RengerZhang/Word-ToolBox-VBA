Attribute VB_Name = "��ʽ_��ݼ�"
Option Explicit

'==========================================================
' ������������������ڶ�/��ѡ���жΡ�ͳһ����ָ����ʽ��д����ʽ���ɵ����ߴ��룩
'==========================================================
Private Sub Ӧ����ʽ_����ѡ����(ByVal ��ʽ�� As String)
    '��һ����ȡѡ������ʽ
    Dim doc As Document: Set doc = ActiveDocument
    Dim s As Style:        Set s = doc.Styles(��ʽ��)
    Dim r As Range:        Set r = Selection.Range
    Dim p As Paragraph
    
    '���������ˢ��ʽ������ѡ�������ȫ�����䣩
    If r.Paragraphs.Count = 0 Then Exit Sub
    For Each p In r.Paragraphs
        p.Range.Style = s
    Next
End Sub

'==========================================================
' ��һ��7 ������꣺����ѡ����ˢ�ɶ�Ӧ��ʽ
'==========================================================
Public Sub һ������():  Ӧ����ʽ_����ѡ���� "���� 1": End Sub
Public Sub ��������():  Ӧ����ʽ_����ѡ���� "���� 2": End Sub
Public Sub ��������():  Ӧ����ʽ_����ѡ���� "���� 3": End Sub
Public Sub �ļ�����():  Ӧ����ʽ_����ѡ���� "���� 4": End Sub
Public Sub �弶����():  Ӧ����ʽ_����ѡ���� "����ʽ��1����": End Sub
Public Sub ��������():  Ӧ����ʽ_����ѡ���� "����ʽ����1����": End Sub
Public Sub �߼�����():  Ӧ����ʽ_����ѡ���� "����ʽ���١�": End Sub
Public Sub ���ĸ�ʽ():  Ӧ����ʽ_����ѡ���� "����": End Sub

'==========================================================
' ��������ݼ���װ/�����Alt+1��Alt+7��Alt+Period������
'   ˵����VBA ��ֱ��ʶ�𡰡�����ֵ��Word ʹ�á�Period(.)����������
'         �����ļ����� Alt+��.�� ����������������ʵ���Ч��
'==========================================================
Public Sub ��װ���������Ŀ�ݼ�_Altϵ��()
    '��1�����浽 Normal��ȫ�֣���������ǰ�ĵ�����Ϊ ActiveDocument
    CustomizationContext = NormalTemplate

    '��2���������λӳ��
    Dim ���� As Variant: ���� = Array( _
        "һ������", "��������", "��������", "�ļ�����", "�弶����", "��������", "�߼�����", _
        "���ĸ�ʽ" _
    )
    
    Dim ���� As Variant
    ���� = Array( _
        BuildKeyCode(AltKeyConst(), wdKey1), _
        BuildKeyCode(AltKeyConst(), wdKey2), _
        BuildKeyCode(AltKeyConst(), wdKey3), _
        BuildKeyCode(AltKeyConst(), wdKey4), _
        BuildKeyCode(AltKeyConst(), wdKey5), _
        BuildKeyCode(AltKeyConst(), wdKey6), _
        BuildKeyCode(AltKeyConst(), wdKey7), _
        BuildKeyCode(AltKeyConst(), wdKeyBackSingleQuote) _
    )   ' Alt + .


    '��3�������ɡ����£������� FindKey���ȣ�
    Dim i As Long
    For i = LBound(����) To UBound(����)
        �����ݼ�_���� CLng(����(i))
        KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                        Command:=CStr(����(i)), _
                        keycode:=CLng(����(i))
    Next

    MsgBox "�Ѱ󶨣�Alt+1~7 ��Ӧһ��~�߼����⣻Alt+����Alt+.��ˢΪ���ġ�", vbInformation
End Sub

Public Sub ������������Ŀ�ݼ�_Altϵ��()
    CustomizationContext = NormalTemplate
    
    Dim ���� As Variant: ���� = Array( _
        BuildKeyCode(AltKeyConst, wdKey1), _
        BuildKeyCode(AltKeyConst, wdKey2), _
        BuildKeyCode(AltKeyConst, wdKey3), _
        BuildKeyCode(AltKeyConst, wdKey4), _
        BuildKeyCode(AltKeyConst, wdKey5), _
        BuildKeyCode(AltKeyConst, wdKey6), _
        BuildKeyCode(AltKeyConst, wdKey7), _
        BuildKeyCode(AltKeyConst, wdKeyBackSingleQuote) _
    )
    
    Dim i As Long
    For i = LBound(����) To UBound(����)
        �����ݼ�_���� CLng(����(i))
    Next
    
    MsgBox "����� Alt+1~7 �� Alt+����Alt+.�����Զ���󶨡�", vbInformation
End Sub

'==========================================================
' ���ģ����ߣ���ƽ̨ Alt ���� & ���������
'==========================================================
Private Function AltKeyConst() As Long
#If Mac Then
    AltKeyConst = wdKeyOption     ' macOS��Option ��
#Else
    AltKeyConst = wdKeyAlt        ' Windows��Alt ��
#End If
End Function

Private Sub �����ݼ�_����(ByVal keycode As Long)
    Dim kb As KeyBinding
    On Error Resume Next
    For Each kb In Application.KeyBindings
        If kb.keycode = keycode Then kb.Clear
    Next
End Sub


