VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 样式_字体样式一键导入 
   Caption         =   "中交标准化样式设置工具"
   ClientHeight    =   9920.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   21030
   OleObjectBlob   =   "样式_字体样式一键导入.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "样式_字体样式一键导入"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' 强制变量声明

Sub ShowMyUserForm()
    UserForm.Show  ' 模态显示窗体
End Sub
Option Explicit
Private Sub btnApply_Click()
    Dim i As Integer
    Dim styleName As String        ' 样式名称
    Dim outlineLevel As Integer    ' 大纲级别（ComboBox 索引）
    Dim outlineLevelVal As Integer ' 实际大纲级别（0~9）
    Dim fontSizeText As String     ' 字号显示文本（如"小三"、"四号"）
    Dim fontSizePt As Single       ' 映射后的字号磅值
    Dim isBold As Boolean          ' 是否加粗
    Dim alignment As Integer       ' 对齐方式（枚举值）
    Dim fontName As String         ' 中文字体
    Dim fontAsciiName As String    ' 西文字体
    Dim indentType As String       ' 缩进类型（首行/悬挂）
    Dim indentValue As Single      ' 缩进值（字符）
    
    ' 遍历 1~10 行样式配置（根据实际需求调整循环范围）
    For i = 1 To 10
        ' ---------------------- 1. 获取控件值（关键：严格匹配控件命名） ----------------------
        ' 样式名称（TextBox01 + 两位行号，如 TextBox0101 ~ TextBox0110）
        styleName = Trim(Me.Controls("TextBox01" & Format(i, "00")).Value)
        If styleName = "" Then
            MsgBox "第 " & i & " 行样式名称为空，请填写后重试！", vbExclamation, "输入错误"
            Exit Sub
        End If
        
        ' 大纲级别（ComboBox02 + 两位行号，如 ComboBox0201 ~ ComboBox0210）
        If Me.Controls("ComboBox02" & Format(i, "00")).ListIndex < 0 Then
            MsgBox "第 " & i & " 行大纲级别未选择，请设置后重试！", vbExclamation, "输入错误"
            Exit Sub
        End If
        outlineLevel = Me.Controls("ComboBox02" & Format(i, "00")).ListIndex
        outlineLevelVal = outlineLevel + 1  ' ListIndex 从 0 开始，实际级别 +1
        
        ' ---------------------- 2. 字号映射处理（核心实现） ----------------------
        ' 获取字号显示文本（如"小三"、"12"）
        fontSizeText = Me.Controls("ComboBox03" & Format(i, "00")).Value
        If fontSizeText = "" Then
            MsgBox "第 " & i & " 行字号未选择，请设置后重试！", vbExclamation
            Exit Sub
        End If
        
        ' 调用映射函数转换为磅值
        fontSizePt = GetFontSizePt(fontSizeText)
        If fontSizePt <= 0 Then
            MsgBox "第 " & i & " 行字号无效：" & fontSizeText, vbExclamation
            Exit Sub
        End If
        
        
        ' 是否加粗（CheckBox04 + 两位行号，如 CheckBox0401 ~ CheckBox0410）
        isBold = (Me.Controls("CheckBox04" & Format(i, "00")).Value = True)
        
        ' 对齐方式（ComboBox05 + 两位行号，如 ComboBox0501 ~ ComboBox0510）
        alignment = Me.Controls("ComboBox05" & Format(i, "00")).ListIndex
        Select Case alignment
            Case 0: alignment = wdAlignParagraphLeft    ' 左对齐
            Case 1: alignment = wdAlignParagraphCenter  ' 居中对齐
            Case 2: alignment = wdAlignParagraphRight   ' 右对齐
            Case 3: alignment = wdAlignParagraphJustify ' 两端对齐
            Case Else: alignment = wdAlignParagraphLeft ' 默认左对齐
        End Select
        
        ' 中文字体（ComboBox06 + 两位行号，如 ComboBox0601 ~ ComboBox0610）
        fontName = Me.Controls("ComboBox06" & Format(i, "00")).Value
        If fontName = "" Then
            MsgBox "第 " & i & " 行中文字体未选择，请设置后重试！", vbExclamation, "输入错误"
            Exit Sub
        End If
        
        ' 西文字体（ComboBox07 + 两位行号，如 ComboBox0701 ~ ComboBox0710）
        fontAsciiName = Me.Controls("ComboBox07" & Format(i, "00")).Value
        If fontAsciiName = "" Then
            MsgBox "第 " & i & " 行西文字体未选择，请设置后重试！", vbExclamation, "输入错误"
            Exit Sub
        End If
        
        ' 缩进类型（ComboBox08 + 两位行号，如 ComboBox0801 ~ ComboBox0810）
        indentType = Me.Controls("ComboBox08" & Format(i, "00")).Value
        Select Case indentType
            Case "首行缩进": indentValue = 2  ' 示例：首行缩进 2 字符
            Case "悬挂缩进": indentValue = -2 ' 示例：悬挂缩进 2 字符（负值表示悬挂）
            Case Else: indentValue = 0        ' 无缩进
        End Select
        
        ' ---------------------- 2. 创建/修改样式（核心逻辑） ----------------------
        Dim myStyle As Style
        On Error Resume Next
        Set myStyle = ActiveDocument.Styles(styleName)  ' 尝试获取已有样式
        On Error GoTo 0
        
        ' 样式不存在则新建，存在则复用
        If myStyle Is Nothing Then
            Set myStyle = ActiveDocument.Styles.Add( _
                name:=styleName, _
                Type:=wdStyleTypeParagraph _
            )
        End If
        
        ' ---------------------- 3. 设置样式属性（逐行配置） ----------------------
        With myStyle
            ' 大纲级别
            .ParagraphFormat.outlineLevel = outlineLevelVal
            
            ' 字体（中/西文）
            .Font.name = fontName          ' 中文字体
            .Font.NameAscii = fontAsciiName ' 西文字体
            .Font.Size = fontSizePt        ' 字号（磅）
            .Font.bold = isBold            ' 是否加粗
            
            ' 段落对齐
            .ParagraphFormat.alignment = alignment
            
            ' 缩进（首行/悬挂）
            .ParagraphFormat.FirstLineIndent = indentValue
            
            ' 可扩展：其他属性（如行距、段前段后间距等）
            '.ParagraphFormat.LineSpacing = 1.5  ' 示例：1.5 倍行距
        End With
        
        ' ---------------------- 4. 提示反馈（可选：告知用户进度） ----------------------
        MsgBox "样式 '" & styleName & "' 创建/修改成功！" & vbCrLf & _
               "→ 大纲级别：" & outlineLevelVal & vbCrLf & _
               "→ 字号：" & fontSizePt & " 磅", vbInformation, "操作完成"
    Next i
    
    MsgBox "所有样式配置已完成！", vbInformation, "批量处理结束"
End Sub


' 用户窗体初始化：为控件提供初始值
Private Sub UserForm_Initialize()
    Dim i As Integer
    For i = 1 To 10
        ' 初始化大纲级别下拉框
        With Me.Controls("ComboBox02" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "无"  ' 无大纲级别选项
            .AddItem "1级"
            .AddItem "2级"
            .AddItem "3级"
            .AddItem "4级"
            .AddItem "5级"
        End With

        ' 初始化字号下拉框
        With Me.Controls("ComboBox03" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "初号": .AddItem "小初": .AddItem "一号": .AddItem "小一"
            .AddItem "二号": .AddItem "小二": .AddItem "三号": .AddItem "小三"
            .AddItem "四号": .AddItem "小四": .AddItem "五号": .AddItem "小五"
            .AddItem "8": .AddItem "9": .AddItem "10": .AddItem "12"
            .AddItem "14": .AddItem "16": .AddItem "18"
        End With
        
        With Me.Controls("ComboBox05" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "左对齐"
            .AddItem "居中对齐"
            .AddItem "右对齐"
            .AddItem "分散对齐"
        End With

        ' 初始化中文字体下拉框
        With Me.Controls("ComboBox06" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "宋体"
            .AddItem "黑体"
            .AddItem "仿宋"
            .AddItem "微软雅黑"
            .AddItem "楷体"
        End With

        ' 初始化西文字体下拉框
        With Me.Controls("ComboBox07" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "Times New Roman"
            .AddItem "Arial"
            .AddItem "Verdana"
        End With

        ' 初始化特殊缩进下拉框
        With Me.Controls("ComboBox08" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "无"
            .AddItem "首行缩进"
            .AddItem "悬挂缩进"
        End With
        
        ' 初始化段前段后间距单位
        With Me.Controls("ComboBox10" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "磅"
            .AddItem "行"
            .AddItem "英寸"
            .AddItem "厘米"
            .AddItem "毫米"
        End With
        
        With Me.Controls("ComboBox12" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "磅"
            .AddItem "行"
            .AddItem "英寸"
            .AddItem "厘米"
            .AddItem "毫米"
        End With
        
        With Me.Controls("ComboBox13" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "单倍行距"
            .AddItem "多倍行距"
            .AddItem "固定值"
            .AddItem "最小值"
        End With

        With Me.Controls("ComboBox15" & Format(i, "00"))
            .Clear ' 清空所有选项
            .AddItem "磅"
            .AddItem "行"
            .AddItem "英寸"
            .AddItem "厘米"
            .AddItem "毫米"
        End With

    Next i
    
End Sub

' 关闭按钮：关闭窗体
Private Sub btnClose_Click()
    Unload Me
End Sub

' ---------- 工具函数：字号文本转磅值 ----------
Private Function GetFontSizePt(sizeText As String) As Single
    Dim sizeMap As Object
    Set sizeMap = CreateObject("Scripting.Dictionary")
    With sizeMap
        .Add "初号", 42
        .Add "小初", 36
        .Add "一号", 26
        .Add "小一", 24
        .Add "二号", 22
        .Add "小二", 18
        .Add "三号", 16
        .Add "小三", 15
        .Add "四号", 14
        .Add "小四", 12
        .Add "五号", 10.5
        .Add "小五", 9
        .Add "六号", 7.5
        .Add "小六", 6.5
    End With

    If sizeMap.exists(sizeText) Then
        GetFontSizePt = sizeMap(sizeText)
    ElseIf IsNumeric(sizeText) Then
        GetFontSizePt = CSng(sizeText)
    Else
        GetFontSizePt = -1
    End If
End Function

' ---------- 工具函数：获取特殊缩进值 ----------
Private Function GetIndentValue(indentType As String) As Single
    If indentType = "首行缩进" Then
        GetIndentValue = 1.5 ' 首行缩进的值
    Else
        GetIndentValue = 0 ' 无缩进
    End If
End Function


