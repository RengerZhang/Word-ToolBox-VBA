Attribute VB_Name = "ȫ��_��ʽ_���δʹ����ʽ"
Sub DeleteUnusedStyles2()
    On Error GoTo ErrorHandler
    Dim oStyle As Style, i&
    i = 0
    For Each oStyle In ActiveDocument.Styles
        'If oStyle.BuiltIn = False Then
            With ActiveDocument.content.Find
                .ClearFormatting
                .MatchWildcards = False
                .Style = CVar(oStyle.NameLocal)
                .Execute findText:="", Format:=True
                If Not .Found Then
                    Application.OrganizerDelete _
                    Source:=ActiveDocument.path & "\" & ActiveDocument.name, _
                    name:=oStyle.NameLocal, Object:=wdOrganizerObjectStyles
                    i = i + 1
                End If
            End With
        'End If
    Next oStyle
MsgBox "��ɾ��" & i & "δʹ����ʽ"
Exit Sub '�˳�����

'��������ʱ����
ErrorHandler:
    i = i - 1 '����һ�δ������1
    Resume Next
End Sub

