Attribute VB_Name = "字号转换函数"
Sub GenerateStyleSettingForm()
    Dim frm As Object
    Set frm = VBA.UserForms.Add("UserForm1")
    
    Dim styleNames As Variant
    styleNames = Array("一级标题", "二级标题", "三级标题", "表标题", "图标题", "表格文字")
    
    Dim colTitles As Variant
    colTitles = Array("样式名称", "大纲级别", "字号", "加粗")
    
    Dim lefts As Variant
    lefts = Array(100, 300, 450, 600) ' 每列控件起始 Left
    
    Dim i As Long, j As Long
    Dim topBase As Long: topBase = 40
    Dim rowHeight As Long: rowHeight = 25

    ' 设置表头
    For j = 0 To UBound(colTitles)
        With frm.Controls.Add("Forms.Label.1", "lblCol" & j, True)
            .caption = colTitles(j)
            .Left = lefts(j)
            .Top = 10
            .width = 100
        End With
    Next j
    
    ' 添加每行样式设置
    For i = 0 To UBound(styleNames)
        Dim topOffset As Long
        topOffset = topBase + i * rowHeight

        ' 行标题
        With frm.Controls.Add("Forms.Label.1", "lblRow" & i, True)
            .caption = styleNames(i)
            .Left = 10
            .Top = topOffset
            .width = 80
        End With
        
        ' 样式名称 TextBox
        With frm.Controls.Add("Forms.TextBox.1", "txtStyleName" & i, True)
            .Left = lefts(0)
            .Top = topOffset
            .width = 100
        End With

        ' 大纲级别 ComboBox
        With frm.Controls.Add("Forms.ComboBox.1", "cmbOutlineLevel" & i, True)
            .Left = lefts(1)
            .Top = topOffset
            .width = 80
            .AddItem "无"
            Dim k
            For k = 1 To 9: .AddItem k & "级": Next k
        End With

        ' 字号 ComboBox
        With frm.Controls.Add("Forms.ComboBox.1", "cmbFontSize" & i, True)
            .Left = lefts(2)
            .Top = topOffset
            .width = 80
            .AddItem "初号": .AddItem "小四": .AddItem "12": .AddItem "10.5"
        End With

        ' 加粗 CheckBox
        With frm.Controls.Add("Forms.CheckBox.1", "chkBold" & i, True)
            .Left = lefts(3)
            .Top = topOffset
            .caption = "加粗"
            .Value = False
        End With
    Next i

    frm.Show
End Sub

