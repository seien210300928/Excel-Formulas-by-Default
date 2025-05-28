# Excel默认值(公式实现)<br>Excel-Formulas-by-Default

当指定范围内的表格单元格为空时，自动填写指定的公式。  
Fill the specified formula when the table cells in the specified range are null.

使用前，请确保电子表格文件为 Microsoft Office 365 Excel 启用宏的工作簿格式（.xlsm）。  
Before using, please ensure that the spreadsheet file is in the format of a Microsoft Office 365 Excel Macro-Enabled Workbook (.xlsm).

对于其他软件或文件格式，甚至是不同版本的Microsoft Office Excel，此 VBA 代码的有效性不保证。  
When you are using other softwares or file formats, or even different versions of Microsoft Office Excel, there won't be working of VBA codes.

<pre><code class="language-vba line-numbers">
' 工作表变更事件处理程序 - 当用户修改工作表中的单元格时触发。
' Worksheet change event handler: This function is triggered when a user modifies any table cells within the worksheet.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' 禁用事件处理，防止代码触发自身形成无限循环。
    ' Disable event handling to prevent the code from triggering itself and forming an infinite loop.
    Application.EnableEvents = False
    ' 检查变更的单元格是否位于目标范围内(Range)。
    ' Check if the changed table cells is within the target range (Range).
    If Not Intersect(Target, Range("<Range_1>")) Is Nothing Then
        ' 遍历所有在目标范围内且被修改的单元格(Range_2=Range_1)。
        ' Loop through all modified table cells within the target range (Range_2 = Range_1).
        For Each cell In Intersect(Target, Range("<Range_2>"))
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
                If cell.<Row,Column> = <Row_value,Column_value> Then
                    ' 指定行或列单元格设置指定公式。
                    ' Set specified formulas for table cells in specified rows or columns.
                    formula = "=<Formula_1>"
                Else
                    ' 非指定行或列单元格设置指定公式。若不需要额外处理指定单元格，则可以删掉此层If_Else语句。
                    ' Set specified formulas for table cells that are not in the specified rows or columns.If no additional processing of specified table cells is required, this layer of If_Else statement can be deleted.
                    formula = "=<Formula_2>")"
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
</code></pre>

## 其中`<>`为替换标签，以下为详细解释
### `<Range_1>`&&`<Range_2>`
`<Range_1>`为代码生效范围；`<Range_2>`为代码执行范围，通常来说这两个值相等。

#### 如果不相等可能会出现以下情况
1. 事件触发条件与处理范围分离  
触发条件：代码仅在用户修改 `Range_1` 内的单元格时触发（通过 `Intersect(Target, Range("<Range_1>"))` 判断）。  
处理范围：但后续遍历的是 `Range_2` 内的单元格（通过 `Intersect(Target, Range("<Range_2>"))`）。  
<br>
可能的问题：  
*如果 `Range_2` 包含 `Range_1` 以外的单元格，这些额外的单元格即使被修改也不会触发事件，导致代码无法处理。
*如果 `Range_2` 是 `Range_1` 的子集，只有子集内的修改会被处理，其他部分会被忽略。  

2. 公式应用逻辑异常  
假设用户修改了 `Range_1` 但未修改 `Range_2` 中的单元格：  
<br>
*代码会遍历 `Range_2` 与 `Target` 的交集，但此时交集可能为空，导致 `For Each` 循环不执行任何操作。  
*即使 `Range_2` 外的单元格被清空，也不会应用公式。

3. 示例场景  
假设：  
`Range_1 = "A1:B10"`（事件触发范围）  
`Range_2 = "C1:D10"`（公式应用范围）  
<br>
当用户修改 `A1:B10` 内的单元格时：  
代码会触发，但 `Intersect(Target, Range("C1:D10"))` 可能为空（因为 `Target` 在 `A1:B10` 内），导致公式无法应用到任何单元格。

4. 潜在风险  
无限循环风险：如果 `Range_2` 与 `Range_1` 有重叠，且公式计算结果可能影响 `Range_1` 内的单元格，可能导致事件被反复触发（即使有 `EnableEvents = False` 保护）。  
逻辑错误：代码注释中提到 `Range_2=Range_1`，说明原设计期望两个范围一致。不一致时违背开发者意图。

### `<Row,Column>`
选择指定行或列。  
*若选择行，则`<Row,Column>`直接替换为`Row`  
*若选择列，则`<Row,Column>`直接替换为`Column`

### `<Row_value,Column_value>`
指定行或列的数值，直接替换为相应数字即可。
若指定的列，则需要将**纯字母27进制**转为**10进制**后替换

### `<Formula_1>`&&`<Formula_2>`
*`<Formula_1>`为特殊处理范围指定公式  
*`<Formula_2>`为非特殊处理范围指定公式


```js:/static/info-logo-script.js showLineNumbers focusBlur
var _t_author = "ninohx96";
// hl-11
var _t_name = "¡HolaH! Blog";
(function () {
  console.log(
    // hl-00
    `%c💎${window._t_name}%c🍉 ${window._t_author}`,
    "color:#fff;background-color:#d6409f;line-height:20px;border-radius: 5px 0 0 5px;padding: 0 8px",
    // hl-99
    "color:#fff;background-color:#3e63dd;line-height:20px;border-radius: 0 5px 5px 0;padding: 0 8px;",
  );
})();
```
