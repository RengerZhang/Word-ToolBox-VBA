Attribute VB_Name = "图片标题_1_样式匹配"
Option Explicit

'========================
' 自检：是否已成功导入两种“表格样式”
' 说明：不改你的文档格式；仅在文档开头临时插入1×1表做属性验证，随后删除
'========================
Public Sub 自检_表格样式导入状态()
    '（一）常量：样式名（过程内可见，避免模块级冲突）
    Const S_NORMAL As String = "标准表格样式"
    Const S_PIC    As String = "图片定位表"

    '（二）准备对象
    Dim doc As Document: Set doc = ActiveDocument
    Dim okNormal As Boolean, okPic As Boolean
    Dim msg As String: msg = "表格样式导入状态：" & vbCrLf

    '（三）存在性 + 类型（必须为“表格样式”）检查
    okNormal = StyleExistsAsTable(doc, S_NORMAL)
    okPic = StyleExistsAsTable(doc, S_PIC)

    msg = msg & " - [" & S_NORMAL & "] " & IIf(okNormal, "已存在（表格样式）", "未找到") & vbCrLf
    msg = msg & " - [" & S_PIC & "] " & IIf(okPic, "已存在（表格样式）", "未找到")

    '（四）对“图片定位表”做一次“实测”：新建临时表→套样式→读取边框/内边距
    If okPic Then
        Dim rng As Range, tb As Table
        Dim passBorders As Boolean, passPadding As Boolean

        Set rng = doc.Range(0, 0)
        rng.Collapse wdCollapseStart

        On Error Resume Next
        doc.UndoRecord.StartCustomRecord "图片定位表-自检"
        On Error GoTo 0

        Set tb = doc.Tables.Add(rng, 1, 1)
        tb.Style = S_PIC

        passBorders = (tb.Borders.enable = False _
                    And tb.Borders(wdBorderTop).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderBottom).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderLeft).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderRight).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderVertical).LineStyle = wdLineStyleNone)

        passPadding = (tb.TopPadding = 0 And tb.BottomPadding = 0 _
                    And tb.LeftPadding = 0 And tb.RightPadding = 0)

        ' 删除临时表
        tb.Range.Delete

        On Error Resume Next
        doc.UndoRecord.EndCustomRecord
        On Error GoTo 0

        msg = msg & vbCrLf & vbCrLf & "图片定位表（实测属性）：" _
            & vbCrLf & " - 框线全关： " & BoolCN(passBorders) _
            & vbCrLf & " - 内边距为0： " & BoolCN(passPadding)
    End If

    '（五）反馈
    MsgBox msg, IIf(okNormal And okPic, vbInformation, vbExclamation), "样式导入自检"
End Sub

'========================
' 演示：把“所选表格”标记为图片表（目测验证最直观）
'========================
Public Sub 一键将所选表格标记为图片表()
    Const S_PIC As String = "图片定位表"
    If Not StyleExistsAsTable(ActiveDocument, S_PIC) Then
        MsgBox "未找到样式【" & S_PIC & "】。请先执行你的“一键导入样式”。", vbExclamation
        Exit Sub
    End If

    If Selection.Information(wdWithInTable) Then
        Selection.Tables(1).Style = S_PIC
        MsgBox "已将所选表格设置为【图片定位表】。请目测：无边框、内边距为0。", vbInformation
    Else
        MsgBox "请先把光标放到要测试的表格里。", vbExclamation
    End If
End Sub

'========================
' 工具：判断某样式是否存在且为“表格样式”
'========================
Private Function StyleExistsAsTable(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    StyleExistsAsTable = (Not st Is Nothing And st.Type = wdStyleTypeTable)
End Function

'========================
' 工具：布尔值汉化
'========================
Private Function BoolCN(ByVal flag As Boolean) As String
    BoolCN = IIf(flag, "是", "否")
End Function


