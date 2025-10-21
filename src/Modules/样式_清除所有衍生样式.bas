Attribute VB_Name = "样式_清除所有衍生样式"
Option Explicit

'========================================
' 一键：全篇仅保留默认样式
'  - 段落样式 → Normal（wdStyleNormal）
'  - 字符样式 → Default Paragraph Font（wdStyleDefaultParagraphFont）
'  - 删除所有非内置的段落/字符/链接样式
' 说明：保留直接格式（加粗/斜体/颜色等）不动
'========================================
Public Sub 还原到默认样式_并清除非默认()
    '（一）准备与性能开关
    Dim doc As Document: Set doc = ActiveDocument
    Dim t0 As Single: t0 = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    '（二）取默认样式对象（区域无关，用内置常量更稳）
    Dim styParagraphDefault As Style
    Dim styCharacterDefault As Style
    Set styParagraphDefault = doc.Styles(wdStyleNormal)                 ' Normal / 正文
    Set styCharacterDefault = doc.Styles(wdStyleDefaultParagraphFont)   ' Default Paragraph Font / 默认段落字体
    
    '（三）A 步：所有故事层 → 段落样式统一还原为 Normal
    Call 故事层_全部设为段落样式(doc, styParagraphDefault)
    
    '（四）B 步：清除所有“非默认”的字符样式（包括内置/自定义/链接样式的字符用法）
    '     思路：对每一种“字符或链接样式”，只要不是 Default Paragraph Font，就用 Find 全文替换为它
    Dim s As Style
    For Each s In doc.Styles
        If s.Type = wdStyleTypeCharacter Or s.Type = wdStyleTypeLinked Then
            If Not (s Is styCharacterDefault) And Not (s Is styParagraphDefault) Then
                Call 故事层_按样式替换为(doc, s, styCharacterDefault)
            End If
        End If
    Next s
    
    '（五）可选：如果连“直接字符格式”也想一起清掉（更干净），放开下一行：
    'doc.Content.ClearCharacterDirectFormatting
    
    '（六）C 步：为避免“样式之间的继承/下一段依赖”阻止删除，先把非内置样式的依赖改到默认
    Call 样式_解除依赖到默认(doc, styParagraphDefault)
    
    '（七）D 步：删除全部“非内置”的 段落/字符/链接 样式（与是否在用已无关）
    Call 删除非内置_段落字符链接样式(doc)
    
    '（八）收尾与提示
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    MsgBox "已完成：仅保留默认样式" & vbCrLf & _
           "· 段落样式 → Normal" & vbCrLf & _
           "· 字符样式 → Default Paragraph Font" & vbCrLf & _
           "· 非内置样式已删除" & vbCrLf & _
           "· 耗时（秒）：" & Format$(Timer - t0, "0.0"), _
           vbInformation, "样式还原完成"
End Sub

'========================================
'（辅助一）把所有故事层的段落样式都设为某段落样式（通常是 Normal）
'========================================
Private Sub 故事层_全部设为段落样式(doc As Document, ByVal paraStyle As Style)
    Dim rng As Range, r2 As Range
    For Each rng In doc.StoryRanges
        ' ——整块应用段落样式（对段落生效；字符样式不会因此被清除）
        rng.Style = paraStyle
        ' ——后续故事层（例如多个文本框串联）
        Set r2 = rng
        Do While Not r2.NextStoryRange Is Nothing
            Set r2 = r2.NextStoryRange
            r2.Style = paraStyle
        Loop
    Next rng
End Sub

'========================================
'（辅助二）在所有故事层，把 oldS → newS（用于字符/链接样式替换）
'========================================
Private Sub 故事层_按样式替换为(doc As Document, ByVal oldS As Style, ByVal newS As Style)
    Dim rng As Range, r2 As Range
    For Each rng In doc.StoryRanges
        Call 范围_按样式替换(rng, oldS, newS)
        Set r2 = rng
        Do While Not r2.NextStoryRange Is Nothing
            Set r2 = r2.NextStoryRange
            Call 范围_按样式替换(r2, oldS, newS)
        Loop
    Next rng
End Sub

Private Sub 范围_按样式替换(ByVal rng As Range, ByVal oldS As Style, ByVal newS As Style)
    With rng.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .Style = oldS
        .replacement.Style = newS
        .Execute Replace:=wdReplaceAll
    End With
End Sub

'========================================
'（辅助三）把所有“非内置”的 段落/字符/链接 样式的基样式/下一段样式 指到 Normal
' 目的：防止样式之间互为 BaseStyle 或 NextParagraphStyle 导致删除失败
'========================================
Private Sub 样式_解除依赖到默认(doc As Document, ByVal paraDefault As Style)
    Dim s As Style
    For Each s In doc.Styles
        On Error Resume Next
        If Not s.BuiltIn Then
            If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeLinked Then
                ' ——将继承链与“下一段样式”改到 Normal
                s.BaseStyle = paraDefault
                s.NextParagraphStyle = paraDefault
            End If
        End If
        On Error GoTo 0
    Next s
End Sub

'========================================
'（辅助四）删除所有“非内置”的 段落/字符/链接 样式
'========================================
Private Sub 删除非内置_段落字符链接样式(doc As Document)
    Dim s As Style
    For Each s In doc.Styles
        If Not s.BuiltIn Then
            Select Case s.Type
                Case wdStyleTypeParagraph, wdStyleTypeCharacter, wdStyleTypeLinked
                    On Error Resume Next
                    s.Delete
                    Err.Clear
                    On Error GoTo 0
            End Select
        End If
    Next s
End Sub


