Attribute VB_Name = "编号_4_自动删除手工编号"
Option Explicit

'==========================================================
' ④ 自动删除手工编号（独立运行）
' 说明：
'   - 目标样式与删除规则：均由 Mod配置中心 动态提供
'   - 仅删除“段首”手工编号，不影响自动编号
'   - 处理完一段后显式跳到下一段，避免最后一段死循环
'==========================================================
Sub 去除手工编号_基于配置中心()

    Dim doc As Document
    Dim backupPath As String
    Dim targetStyles As Variant ' 样式名数组（只取文档中存在的）
    Dim patterns As Variant     ' 删除规则数组（基于编号格式动态生成）
    
    Dim styleName As Variant
    Dim rng As Range, contentRng As Range
    Dim originalText As String, newText As String
    Dim pat As Variant
    Dim matched As Boolean
    
    Dim lastStart As Long
    Dim nextPos As Long
    
    Set doc = ActiveDocument
    
    '——运行前自动备份（同目录）
    backupPath = 备份当前文档(doc)
    If Len(backupPath) > 0 Then Debug.Print "已备份到: " & backupPath
    
    '——从配置中心取样式 & 规则（不再手填）
    targetStyles = 获取样式名数组(True)    ' 只返回文档中存在的样式
    patterns = 生成删除编号规则集()         ' 基于（③）的编号格式自动推断
    
    '——遍历每一种目标样式（按样式 Find，精准且高效）
    For Each styleName In targetStyles
        
        Set rng = doc.content
        lastStart = -1
        
        With rng.Find
            .ClearFormatting
            .Style = doc.Styles(CStr(styleName))
            .text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            
            Do While .Execute
                '——死循环防护（若起点未推进，则强制跳下一段）
                If rng.Start = lastStart Then
                    nextPos = rng.Paragraphs(1).Range.End
                    If nextPos >= doc.content.End Then Exit Do
                    rng.SetRange Start:=nextPos, End:=doc.content.End
                    lastStart = -1
                    GoTo ContinueFindLoop
                End If
                lastStart = rng.Start
                
                '——复制内容子范围：不含段尾标记（?）
                Set contentRng = rng.Duplicate
                If contentRng.Characters.Count > 1 Then
                    contentRng.MoveEnd wdCharacter, -1
                End If
                
                originalText = contentRng.text
                newText = originalText
                matched = False
                
                '——单次 pass：把所有删除规则依次跑一遍（仅段首；每条只替一次）
                For Each pat In patterns
                    If 正则命中(newText, CStr(pat)) Then
                        newText = 正则替换(newText, CStr(pat), "")
                        matched = True
                    End If
                Next
                
                '——如有命中，清理段首残留空格（含全角）
                If matched Then
                    newText = 正则替换(newText, "^[ 　]+", "")
                End If
                
                '——仅在有变化时写回
                If matched And newText <> originalText Then
                    contentRng.text = newText
                End If
                
                '——关键：显式跳到下一段，彻底杜绝末段死循环
                nextPos = rng.Paragraphs(1).Range.End
                If nextPos >= doc.content.End Then Exit Do
                rng.SetRange Start:=nextPos, End:=doc.content.End
                
ContinueFindLoop:
            Loop
        End With
    Next styleName
    
    MsgBox "④ 删除手工编号完成（样式/规则均自动读取配置）。", vbInformation
End Sub

'——简易正则封装（与②一致）
Private Function 正则替换(ByVal s As String, ByVal pat As String, ByVal rep As String) As String
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.pattern = pat
    正则替换 = r.Replace(s, rep)
End Function

Private Function 正则命中(ByVal s As String, ByVal pat As String) As Boolean
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.pattern = pat
    正则命中 = r.TEST(s)
End Function

'——备份（同你现有的函数一致，可复用）
Private Function 备份当前文档(ByVal doc As Document) As String
    On Error GoTo EH

    Dim baseName As String, ext As String, bak As String
    Dim folder As String, ts As String

    ts = Format(Now, "yyyymmdd_hhnnss")

    If Len(doc.name) > 0 Then
        baseName = doc.name
        If InStrRev(baseName, ".") > 0 Then
            ext = mid$(baseName, InStrRev(baseName, "."))
            baseName = Left$(baseName, InStrRev(doc.name, ".") - 1)
        Else
            ext = ".docx"
        End If
    Else
        baseName = "未命名文档"
        ext = ".docx"
    End If

    folder = IIf(doc.path = "", CurDir$, doc.path)
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    bak = folder & baseName & "_备份_" & ts & ext
    doc.SaveCopyAs FileName:=bak

    备份当前文档 = bak
    Exit Function

EH:
    备份当前文档 = ""
End Function

