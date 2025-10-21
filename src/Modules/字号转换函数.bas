Attribute VB_Name = "�ֺ�ת������"
Sub GenerateStyleSettingForm()
    Dim frm As Object
    Set frm = VBA.UserForms.Add("UserForm1")
    
    Dim styleNames As Variant
    styleNames = Array("һ������", "��������", "��������", "�����", "ͼ����", "�������")
    
    Dim colTitles As Variant
    colTitles = Array("��ʽ����", "��ټ���", "�ֺ�", "�Ӵ�")
    
    Dim lefts As Variant
    lefts = Array(100, 300, 450, 600) ' ÿ�пؼ���ʼ Left
    
    Dim i As Long, j As Long
    Dim topBase As Long: topBase = 40
    Dim rowHeight As Long: rowHeight = 25

    ' ���ñ�ͷ
    For j = 0 To UBound(colTitles)
        With frm.Controls.Add("Forms.Label.1", "lblCol" & j, True)
            .caption = colTitles(j)
            .Left = lefts(j)
            .Top = 10
            .width = 100
        End With
    Next j
    
    ' ���ÿ����ʽ����
    For i = 0 To UBound(styleNames)
        Dim topOffset As Long
        topOffset = topBase + i * rowHeight

        ' �б���
        With frm.Controls.Add("Forms.Label.1", "lblRow" & i, True)
            .caption = styleNames(i)
            .Left = 10
            .Top = topOffset
            .width = 80
        End With
        
        ' ��ʽ���� TextBox
        With frm.Controls.Add("Forms.TextBox.1", "txtStyleName" & i, True)
            .Left = lefts(0)
            .Top = topOffset
            .width = 100
        End With

        ' ��ټ��� ComboBox
        With frm.Controls.Add("Forms.ComboBox.1", "cmbOutlineLevel" & i, True)
            .Left = lefts(1)
            .Top = topOffset
            .width = 80
            .AddItem "��"
            Dim k
            For k = 1 To 9: .AddItem k & "��": Next k
        End With

        ' �ֺ� ComboBox
        With frm.Controls.Add("Forms.ComboBox.1", "cmbFontSize" & i, True)
            .Left = lefts(2)
            .Top = topOffset
            .width = 80
            .AddItem "����": .AddItem "С��": .AddItem "12": .AddItem "10.5"
        End With

        ' �Ӵ� CheckBox
        With frm.Controls.Add("Forms.CheckBox.1", "chkBold" & i, True)
            .Left = lefts(3)
            .Top = topOffset
            .caption = "�Ӵ�"
            .Value = False
        End With
    Next i

    frm.Show
End Sub

