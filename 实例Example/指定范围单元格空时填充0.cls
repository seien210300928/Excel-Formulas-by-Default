VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 工作表变更事件处理程序 - 当用户修改工作表中的单元格时触发
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' 禁用事件处理，防止代码触发自身形成无限循环
    Application.EnableEvents = False
    
    ' 检查变更的单元格是否位于D7:AA8范围内
    If Not Intersect(Target, Range("D7:AA8")) Is Nothing Then
        ' 遍历所有在D7:AA8范围内且被修改的单元格
        For Each cell In Intersect(Target, Range("D7:AA8"))
            ' 如果单元格被清空（值为空字符串）
            If cell.Value = "" Then
                ' 声明变量存储列字母（如D、E、F等）
                Dim colLetter As String
                ' 通过拆分单元格地址获取列字母部分（例如$E$7拆分为"E"）
                colLetter = Split(cell.Address, "$")(1)
                ' 声明变量存储要设置的公式
                Dim formula As String
                ' 构建公式逻辑：
                ' 检查当前单元格所在行的内容是否为空
                ' 如果为空则显示为空字符串，否则显示0
                formula = "=IF(ISBLANK(" & colLetter & (cell.Row) & "), """", 0)"
                ' 例如D7会生成公式: =IF(ISBLANK(D7), "", 0)
                ' 将构建好的公式设置到单元格中
                cell.formula = formula
            End If
        Next cell
    End If
    
    ' 恢复事件处理
    Application.EnableEvents = True
End Sub
