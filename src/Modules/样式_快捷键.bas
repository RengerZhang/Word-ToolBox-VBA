Attribute VB_Name = "样式_快捷键"
Option Explicit

'==========================================================
' （公共）将“光标所在段/所选所有段”统一套用指定样式（写死样式名由调用者传入）
'==========================================================
Private Sub 应用样式_到所选段落(ByVal 样式名 As String)
    '（一）获取选区与样式
    Dim doc As Document: Set doc = ActiveDocument
    Dim s As Style:        Set s = doc.Styles(样式名)
    Dim r As Range:        Set r = Selection.Range
    Dim p As Paragraph
    
    '（二）逐段刷样式（覆盖选区触达的全部段落）
    If r.Paragraphs.Count = 0 Then Exit Sub
    For Each p In r.Paragraphs
        p.Range.Style = s
    Next
End Sub

'==========================================================
' （一）7 个标题宏：把所选段落刷成对应样式
'==========================================================
Public Sub 一级标题():  应用样式_到所选段落 "标题 1": End Sub
Public Sub 二级标题():  应用样式_到所选段落 "标题 2": End Sub
Public Sub 三级标题():  应用样式_到所选段落 "标题 3": End Sub
Public Sub 四级标题():  应用样式_到所选段落 "标题 4": End Sub
Public Sub 五级标题():  应用样式_到所选段落 "条样式【1）】": End Sub
Public Sub 六级标题():  应用样式_到所选段落 "款样式【（1）】": End Sub
Public Sub 七级标题():  应用样式_到所选段落 "项样式【①】": End Sub
Public Sub 正文格式():  应用样式_到所选段落 "正文": End Sub

'==========================================================
' （三）快捷键安装/清除：Alt+1…Alt+7；Alt+Period（·）
'   说明：VBA 不直接识别“·”键值，Word 使用“Period(.)”键常量；
'         在中文键盘上 Alt+“.” 往往产生“·”，实测等效。
'==========================================================
Public Sub 安装标题与正文快捷键_Alt系列()
    '（1）保存到 Normal（全局）；若仅当前文档，改为 ActiveDocument
    CustomizationContext = NormalTemplate

    '（2）宏名与键位映射
    Dim 宏名 As Variant: 宏名 = Array( _
        "一级标题", "二级标题", "三级标题", "四级标题", "五级标题", "六级标题", "七级标题", _
        "正文格式" _
    )
    
    Dim 键码 As Variant
    键码 = Array( _
        BuildKeyCode(AltKeyConst(), wdKey1), _
        BuildKeyCode(AltKeyConst(), wdKey2), _
        BuildKeyCode(AltKeyConst(), wdKey3), _
        BuildKeyCode(AltKeyConst(), wdKey4), _
        BuildKeyCode(AltKeyConst(), wdKey5), _
        BuildKeyCode(AltKeyConst(), wdKey6), _
        BuildKeyCode(AltKeyConst(), wdKey7), _
        BuildKeyCode(AltKeyConst(), wdKeyBackSingleQuote) _
    )   ' Alt + .


    '（3）逐个清旧→加新（不依赖 FindKey，稳）
    Dim i As Long
    For i = LBound(宏名) To UBound(宏名)
        清除快捷键_遍历 CLng(键码(i))
        KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                        Command:=CStr(宏名(i)), _
                        keycode:=CLng(键码(i))
    Next

    MsgBox "已绑定：Alt+1~7 对应一级~七级标题；Alt+·（Alt+.）刷为正文。", vbInformation
End Sub

Public Sub 清除标题与正文快捷键_Alt系列()
    CustomizationContext = NormalTemplate
    
    Dim 键码 As Variant: 键码 = Array( _
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
    For i = LBound(键码) To UBound(键码)
        清除快捷键_遍历 CLng(键码(i))
    Next
    
    MsgBox "已清除 Alt+1~7 与 Alt+·（Alt+.）的自定义绑定。", vbInformation
End Sub

'==========================================================
' （四）工具：跨平台 Alt 常量 & 遍历清理绑定
'==========================================================
Private Function AltKeyConst() As Long
#If Mac Then
    AltKeyConst = wdKeyOption     ' macOS：Option 键
#Else
    AltKeyConst = wdKeyAlt        ' Windows：Alt 键
#End If
End Function

Private Sub 清除快捷键_遍历(ByVal keycode As Long)
    Dim kb As KeyBinding
    On Error Resume Next
    For Each kb In Application.KeyBindings
        If kb.keycode = keycode Then kb.Clear
    Next
End Sub


