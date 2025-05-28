VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 工作表变更事件处理程序 - 当用户修改工作表中的单元格时触发。
' Worksheet change event handler: This function is triggered when a user modifies any table cells within the worksheet.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' 禁用事件处理，防止代码触发自身形成无限循环。
    ' Disable event handling to prevent the code from triggering itself and forming an infinite loop.
    Application.EnableEvents = False
    ' 检查变更的单元格是否位于目标范围内(Range)。
    ' Check if the changed table cells is within the target range (Range).
    If Not Intersect(Target, Range("E3:E14,I3:I14")) Is Nothing Then
        ' 遍历所有在目标范围内且被修改的单元格(Range_2=Range_1)。
        ' Loop through all modified table cells within the target range (Range_2 = Range_1).
        For Each cell In Intersect(Target, Range("E3:E14,I3:I14"))
            ' 如果单元格被清空（值为空字符串）。
            ' If the table cells is cleared (value is null).
            If cell.Value = "" Then
                ' 声明变量存储列字母（如E或I）。
                ' Declare a variable to store column letters (such as E or I).
                Dim colLetter As String
                ' 通过拆分单元格地址获取列字母部分（例如$E$4拆分为"E"），若指定公式要引用单元格地址则需要此行。
                ' Get the column letter part by splitting the table cells address (e.g., $E$4 is split into "E"). This line is required if the specified formula needs to reference the table cells address.
                colLetter = Split(cell.Address, "$")(1)
                ' 声明变量存储要设置的公式。
                ' Declare a variable to store the formula to be set.
                Dim formula As String
                ' 特殊处理指定行或列的单元格（Row_value,Column_value）.
                ' Special handling for table cells in specified rows or columns (Row_value, Column_value).
                If cell.Row = 3 Then
                    ' 指定行或列单元格设置指定公式。
                    ' Set specified formulas for table cells in specified rows or columns.
                    formula = "=0"
                Else
                    ' 非指定行或列单元格设置指定公式。若不需要额外处理指定单元格，则可以删掉此层If_Else语句。
                    ' Set specified formulas for table cells that are not in the specified rows or columns.If no additional processing of specified table cells is required, this layer of If_Else statement can be deleted.
                    formula = "=IF(ISBLANK(" & colLetter & (cell.Row - 1) & "), """", " & colLetter & (cell.Row - 1) & ")"
                End If
                ' 将构建好的公式设置到单元格中。
                ' Set the constructed formula into the table cells.
                cell.formula = formula
            End If
        Next cell
    End If
    ' 恢复事件处理
    ' Restore event handling
    Application.EnableEvents = True
End Sub
