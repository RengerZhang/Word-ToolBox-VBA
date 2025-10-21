Attribute VB_Name = "编号_3_多级模板定义和匹配"
'==============================
' 功能：为 Word 文档自动建立多级标题编号
' 支持 1~7 级：标题、条、款、项 等样式
'==============================

Sub 标题自动编号()
    '==============================
    ' 主控程序：核心逻辑分 3 步
    '==============================
    Dim 文档 As Document
    Dim 多级模板 As ListTemplate
    Dim 级别参数() As Variant    ' 存储所有级别参数（从函数获取）
    Dim 当前级别 As Integer      ' 循环变量（控制 1-N 级）

    '--- 1. 初始化文档与多级列表模板 ---
    Set 文档 = ActiveDocument
    Set 多级模板 = 文档.ListTemplates.Add(OutlineNumbered:=True)  ' 新建多级模板

    '--- 2. 获取所有级别参数 ---
    '   所有参数集中定义在函数【获取所有级别参数】里
    '   以后要改编号样式、添加新级别，只需要改那个函数
    级别参数 = 获取所有级别参数()

    '--- 3. 核心循环 ---
    '   遍历每个级别，调用【配置单级列表】函数绑定样式
    For 当前级别 = 1 To UBound(级别参数, 1)   ' 自动识别最大级别
        Call 配置单级列表( _
            多级模板.ListLevels(当前级别), _
            文档, _
            级别参数(当前级别, 1), _
            级别参数(当前级别, 2), _
            级别参数(当前级别, 3), _
            级别参数(当前级别, 4) _
        )
    Next 当前级别

    '--- 提示完成 ---
    MsgBox "已完成 " & UBound(级别参数, 1) & " 级编号模板创建，样式绑定成功！", vbInformation
End Sub


'==============================
' 函数：获取所有级别参数
' 返回二维数组：(级别, 参数列)
' 参数列顺序 = 【样式名, 编号格式, 编号样式, 对齐位置(cm)】
'==============================
Public Function 获取所有级别参数() As Variant
    Dim 所有参数(1 To 7, 1 To 4) As Variant   ' 1~7 级（如需更多，扩展此数组）

    '--- 级别1：标题1 ---
    所有参数(1, 1) = "标题 1"
    所有参数(1, 2) = "%1  "
    所有参数(1, 3) = wdListNumberStyleArabic
    所有参数(1, 4) = 0

    '--- 级别2：标题2 ---
    所有参数(2, 1) = "标题 2"
    所有参数(2, 2) = "%1.%2  "
    所有参数(2, 3) = wdListNumberStyleArabic
    所有参数(2, 4) = 0

    '--- 级别3：标题3 ---
    所有参数(3, 1) = "标题 3"
    所有参数(3, 2) = "%1.%2.%3  "
    所有参数(3, 3) = wdListNumberStyleArabic
    所有参数(3, 4) = 0

    '--- 级别4：标题4 ---
    所有参数(4, 1) = "标题 4"
    所有参数(4, 2) = "%1.%2.%3.%4  "
    所有参数(4, 3) = wdListNumberStyleArabic
    所有参数(4, 4) = 0

    '--- 级别5：条（示例：1））---
    所有参数(5, 1) = "条样式【1）】"
    所有参数(5, 2) = "%5）"
    所有参数(5, 3) = wdListNumberStyleArabic
    所有参数(5, 4) = 0

    '--- 级别6：款（示例：（1））---
    所有参数(6, 1) = "款样式【（1）】"
    所有参数(6, 2) = "（%6）"
    所有参数(6, 3) = wdListNumberStyleArabic
    所有参数(6, 4) = 0

    '--- 级别7：项（示例：①）---
    所有参数(7, 1) = "项样式【①】"
    所有参数(7, 2) = "%7 "
    所有参数(7, 3) = wdListNumberStyleNumberInCircle
    所有参数(7, 4) = 0

    '--- 扩展示例 ---
    ' 所有参数(8, 1) = "新增样式名"
    ' 所有参数(8, 2) = "%8、  "
    ' 所有参数(8, 3) = wdListNumberStyleArabic
    ' 所有参数(8, 4) = 1.2

    获取所有级别参数 = 所有参数
End Function


'==============================
' 子过程：配置单个级别的编号规则
' 参数：
'   单级列表  —— 当前级别的 ListLevel 对象
'   文档      —— 当前文档对象
'   样式名    —— Word 样式名（如“标题 1”）
'   编号格式  —— 显示格式（如 "%1.%2.%3"）
'   编号样式  —— 数字样式（阿拉伯、带圈等）
'   对齐位置  —— 编号左对齐位置（单位：cm）
'==============================
Private Sub 配置单级列表( _
    ByRef 单级列表 As ListLevel, _
    ByRef 文档 As Document, _
    ByVal 样式名 As String, _
    ByVal 编号格式 As String, _
    ByVal 编号样式 As WdListNumberStyle, _
    ByVal 对齐位置cm As Single _
)
    With 单级列表
        '--- 所有级别的通用固定配置 ---
        .TrailingCharacter = wdTrailingNone             ' 编号后无特殊标注
        .startAt = 1                                    ' 编号从1开始
        .alignment = wdListLevelAlignLeft               ' 编号左对齐
        .TabPosition = 0                                ' 清除制表位
        .ResetOnHigher = .Index - 1                     ' 上级变化时自动重置

        '--- 每个级别不同的动态配置 ---
        .NumberStyle = 编号样式                         ' 编号样式（阿拉伯/带圈等）
        .NumberFormat = 编号格式                        ' 显示格式（如 1.1/（1））
        .NumberPosition = CentimetersToPoints(对齐位置cm) ' 对齐位置（cm 转磅）
        .TextPosition = .NumberPosition                 ' 文本起始位置
        .LinkedStyle = 文档.Styles(样式名).nameLocal    ' 绑定 Word 样式
    End With
End Sub



