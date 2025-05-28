# Excel默认值(公式实现)  
Excel-Formulas-by-Default

当指定范围内的表格单元格为空时，自动填写指定的公式。  
Fill the specified formula when the table cells in the specified range are null.

使用前，请确保电子表格文件为 Microsoft Office 365 Excel 启用宏的工作簿格式（.xlsm）。  
Before using, please ensure that the spreadsheet file is in the format of a Microsoft Office 365 Excel Macro-Enabled Workbook (.xlsm).

对于其他软件或文件格式，甚至是不同版本的Microsoft Office Excel，此 VBA 代码的有效性不保证。  
When you are using other softwares or file formats, or even different versions of Microsoft Office Excel, there won't be working of VBA codes.

···
' 工作表变更事件处理程序 - 当用户修改工作表中的单元格时触发
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' 禁用事件处理，防止代码触发自身形成无限循环
    Application.EnableEvents = False
    
    ' 检查变更的单元格是否位于目标范围内（E3:E14和I3:I14）
    If Not Intersect(Target, Range("E3:E14,I3:I14")) Is Nothing Then
        ' 遍历所有在目标范围内且被修改的单元格
        For Each cell In Intersect(Target, Range("E3:E14,I3:I14"))
            ' 如果单元格被清空（值为空字符串）
            If cell.Value = "" Then
                ' 声明变量存储列字母（如E或I）
                Dim colLetter As String
                ' 通过拆分单元格地址获取列字母部分（例如$E$4拆分为"E"）
                colLetter = Split(cell.Address, "$")(1)
                
                ' 声明变量存储要设置的公式
                Dim formula As String
                
                ' 特殊处理第3行的单元格（E3和I3）
                If cell.Row = 3 Then
                    ' 第3行单元格直接设置公式为固定值0
                    formula = "=0"
                Else
                    ' 非第3行单元格设置公式引用上一行
                    formula = "=IF(ISBLANK(" & colLetter & (cell.Row - 1) & "), """", " & colLetter & (cell.Row - 1) & ")"
                    ' 例如E4会生成公式: =IF(ISBLANK(E3), "", E3)
                End If
                
                ' 将构建好的公式设置到单元格中
                cell.formula = formula
            End If
        Next cell
    End If
    
    ' 恢复事件处理
    Application.EnableEvents = True
End Sub



···
