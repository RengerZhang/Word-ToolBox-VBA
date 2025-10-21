Attribute VB_Name = "ģ��2"
Option Explicit

'==========================================================
' �������ԣ�����������⡱�ε��ֹ����ǰ׺��������
' ����
'  - Step A���������ԡ�����ͷ��ɾ���ӡ���������һ�������ַ���Ϊֹ
'  - Step B���� A δ���У����á����ݵ���ʧ�ء�������������������̬��
'            �� [��ѡ���ַ�] ���� [�������]* [��ѡ - ����]
'            ����ȫ/��ǵ㣨. �� ������������ַ���- �� �C ����
' ���������ģ�wdMainTextStory��
' Ĭ�Ͻ�������ʽ�������⡿������ʽ���������˻�Ϊ���������ԡ�����ͷ�Ķ��䡱
'==========================================================
Sub ����_��������ֹ����_����()
    Const ������ָ����ʽ As Boolean = True
    Const ������ʽ�� As String = "������"
    
    Dim doc As Document: Set doc = ActiveDocument
    Dim capStyle As Style, useStyleFilter As Boolean
    Dim p As Paragraph, r As Range
    Dim oldTxt As String, newTxt As String
    Dim total As Long, touched As Long, skipped As Long, examples As Long
    
    '������ʽ���ˣ���������Ŀ����ʽ�����˻�Ϊ����������ʽ��
    useStyleFilter = ������ָ����ʽ
    If useStyleFilter Then
        On Error Resume Next
        Set capStyle = doc.Styles(������ʽ��)
        On Error GoTo 0
        If capStyle Is Nothing Then useStyleFilter = False
    End If
    
    '������ѡ���������޸ķ���һ�������飨�°汾 Word ֧�֣�
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "��������ֹ����"
    On Error GoTo 0
    
    For Each p In doc.Paragraphs
        ' ������
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextPara
        
        ' ��ʽ���ˣ�������ã�
        If useStyleFilter Then
            On Error Resume Next
            If p.Range.Style.nameLocal <> ������ʽ�� Then GoTo NextPara
            On Error GoTo 0
        End If
        
        ' ֻ�����ԡ�����ͷ���Ķ��䣨��ǰ���ո�/ȫ�ǿո�
        oldTxt = ������׿ɼ��ı�(p.Range.text)
        If Len(oldTxt) = 0 Or Left$(oldTxt, 1) <> "��" Then GoTo NextPara
        
        '����Ŀ���ӷ�Χ��������β��ǣ�
        Set r = p.Range.Duplicate
        If r.Characters.Count > 1 Then r.MoveEnd wdCharacter, -1
        
        '����Step A���� �� ɾ����һ�������ַ�
        newTxt = ȥ�������ǰ׺_����һ������(r.text)
        
        '����Step B���� A δ�ı��ı����������򶵵ס���[��ѡ���ַ�]����[���]*[��ѡ-����]��
        '   ���ͱ�����
        '   ^\s*                 ���� �Ӷ��׿�ʼ���������ɿհף���ȫ�ǿո���������תΪ��ǣ�
        '   ��\s*                ���� �������������ɿո�
        '   [-���C��]?             ���� ��ѡ���ַ������ǳ����� -��ȫ�ǣ����̺�C�����ᡪ��
        '   \s*\d+               ���� ���ɿո������һλ���֣��ϸ��ֹ����A-1������ƥ�䣩
        '   (?:\s*[\.����]\s*\d+)* ���� 0~��Ρ��� + ���֡������Ϊ���.��ȫ�ǣ������ġ���
        '   \s*                  ���� ��ѡ�ո�
        '   (?:[-���C��]\s*\d+)?   ���� ��ѡ�����ַ� + ���֡���˳��ţ��� -1��
        '   \s*                  ���� �Ե���ź���Ŀո�
        If newTxt = r.text Then
            newTxt = �����滻( _
                newTxt, _
                "^\s*��\s*[-���C��]?\s*\d+(?:\s*[\.����]\s*\d+)*\s*(?:[-���C��]\s*\d+)?\s*", _
                "" _
            )
            newTxt = LTrim$(newTxt)
        End If
        
        total = total + 1
        If newTxt <> r.text Then
            r.text = newTxt
            touched = touched + 1
            ' ��ӡǰ 8 ���޸�ʾ�������������ڡ�
            If examples < 8 Then
                Debug.Print "���ǰ��"; oldTxt
                Debug.Print " �ĺ�"; ������׿ɼ��ı�(newTxt)
                examples = examples + 1
            End If
        Else
            skipped = skipped + 1
        End If
        
NextPara:
    Next p
    
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    
    MsgBox "������ɣ�" & vbCrLf & _
           "��ѡ���䣨�ԡ�����ͷ����" & total & vbCrLf & _
           "�����ǰ׺��" & touched & vbCrLf & _
           "δ�������ƥ�䣩��" & skipped & vbCrLf & vbCrLf & _
           "��ʾ���� Ctrl+G �򿪡��������ڡ��ɲ鿴ʾ����", vbInformation
End Sub

'����Step A�����ԡ�����ͷ���ӡ���ɾ������һ�������ַ���
Private Function ȥ�������ǰ׺_����һ������(ByVal s As String) As String
    Dim i As Long, ch As String, hit As Boolean
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = LTrim$(s)
    
    If Len(s) = 0 Or Left$(s, 1) <> "��" Then
        ȥ�������ǰ׺_����һ������ = s
        Exit Function
    End If
    
    For i = 2 To Len(s)
        ch = mid$(s, i, 1)
        If �Ƿ������ַ�(ch) Then hit = True: Exit For
    Next i
    
    If hit Then
        ȥ�������ǰ׺_����һ������ = LTrim$(mid$(s, i))
    Else
        ȥ�������ǰ׺_����һ������ = s   ' û�ҵ����ģ��������򶵵�
    End If
End Function

'�����ж��Ƿ����ģ����� AscW �������⣻CJK ������ + ��չA��
Private Function �Ƿ������ַ�(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) = 0 Then �Ƿ������ַ� = False: Exit Function
    code = AscW(ch)
    If code < 0 Then code = code + &H10000  ' ��ؼ������з���ֵ��һ�� 0..65535
    �Ƿ������ַ� = ((code >= &H4E00 And code <= &H9FFF) Or (code >= &H3400 And code <= &H4DBF))
End Function

'�������򣺵����滻����Сд�����У�
Private Function �����滻(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim r As Object: Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False          ' ֻ�滻���׵���һ��ǰ׺�������ɡ�������С����Դ���
    r.pattern = pat
    �����滻 = r.Replace(s, rep)
End Function

'������ϴ��ȥ��β��ȥ��Ԫ���������ȫ�ǿո����ǲ� Trim
Private Function ������׿ɼ��ı�(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    ������׿ɼ��ı� = Trim$(s)
End Function


