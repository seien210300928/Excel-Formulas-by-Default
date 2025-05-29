## Language
- [中文](#中文)
- [English](#english)
- [日本語](#日本語)

---
### 中文
# Excel默认值(公式实现)

当指定范围内的表格单元格为空时，自动填写指定的公式。  

1. 使用前，请确保电子表格文件为Microsoft 365MSO(版本 2504 Build 16.0.18730.20186) 64位 Microsoft Excel 启用宏的工作簿格式（.xlsm）。  
2. 对于其他软件或文件格式，甚至是不同版本的Microsoft Office Excel，此 VBA 代码的有效性不保证。  

<pre><code class="language-vba line-numbers">
' 工作表变更事件处理程序 - 当用户修改工作表中的单元格时触发。
modifies any table cells within the worksheet.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' 禁用事件处理，防止代码触发自身形成无限循环。
    forming an infinite loop.
    Application.EnableEvents = False
    ' 检查变更的单元格是否位于目标范围内(Range)。
    If Not Intersect(Target, Range("<Range_1>")) Is Nothing Then
        ' 遍历所有在目标范围内且被修改的单元格(Range_2=Range_1)。
        (Range_2 = Range_1).
        For Each cell In Intersect(Target, Range("<Range_2>"))
            ' 如果单元格被清空（值为空字符串）。
            If cell.Value = "" Then
                ' 声明变量存储列字母（如E或I）。
                Dim colLetter As String
                ' 通过拆分单元格地址获取列字母部分（例如$E$4拆分为"E"），若指定公式要引用单元格地址则需要此行。
                colLetter = Split(cell.Address, "$")(1)
                ' 声明变量存储要设置的公式。
                Dim formula As String
                ' 特殊处理指定行或列的单元格（Row_value,Column_value）.
                columns (Row_value, Column_value).
                If cell.<Row,Column> = <Row_value,Column_value> Then
                    ' 指定行或列单元格设置指定公式。
                    formula = "=<Formula_1>"
                Else
                    ' 非指定行或列单元格设置指定公式。若不需要额外处理指定单元格，则可以删掉此层If_Else语句。
                    formula = "=<Formula_2>")"
                End If
                ' 将构建好的公式设置到单元格中。
            End If
        Next cell
    End If
    ' 恢复事件处理
    Application.EnableEvents = True
End Sub
</code></pre>

## 其中`<>`为替换标签，以下为详细解释：

### `<Range_1>`&&`<Range_2>`
`<Range_1>`为代码生效范围；`<Range_2>`为代码执行范围，通常来说这**两个值相等**。  

#### 如果不相等可能会出现以下情况：

1. ##### 事件触发条件与处理范围分离
    触发条件：代码仅在用户修改`<Range_1>`内的单元格时触发（通过`Intersect(Target, Range("<Range_1>"))`判断）。  
    处理范围：但后续遍历的是`<Range_2>`内的单元格（通过`Intersect(Target, Range("<Range_2>"))`）。  
    <br>
    可能的问题：  
    * 如果`<Range_2>`包含`<Range_1>`以外的单元格，这些额外的单元格即使被修改也不会触发事件，导致代码无法处理。  
    * 如果`<Range_2>`是`<Range_1>`的子集，只有子集内的修改会被处理，其他部分会被忽略。

2. ##### 公式应用逻辑异常
    假设用户修改了`<Range_1>`但未修改`<Range_2>`中的单元格：  
    * 代码会遍历`<Range_2>`与`<Target>`的交集，但此时交集可能为空，导致`<For Each>`循环不执行任何操作。  
    * 即使`<Range_2>`外的单元格被清空，也不会应用公式。  

3. ##### 示例场景
    假设：  
    * `Range_1 = "A1:B10"`（事件触发范围）  
    * `Range_2 = "C1:D10"`（公式应用范围）  
    
    当用户修改`A1:B10`内的单元格时：  
    * 代码会触发，但`Intersect(Target, Range("C1:D10"))`可能为空（因为`Target` 在 `A1:B10`内），导致公式无法应用到任何单元格。  

4. ##### 潜在风险
    无限循环风险：如果`<Range_2>`与`<Range_1>`有重叠，且公式计算结果可能影响`<Range_1>`内的单元格，可能导致事件被反复触发（即使有`EnableEvents = False`保护）。  


### `<Row,Column>`
选择指定行或列。  
* 若选择行，则`<Row,Column>`直接替换为`Row`<br>
* 若选择列，则`<Row,Column>`直接替换为`Column`<br>

### `<Row_value,Column_value>`
指定行或列的数值，直接替换为相应数字即可。   
若指定的列，则需要将**纯字母27进制**转为**10进制**后替换。  

### `<Formula_1>`&&`<Formula_2>`
* `<Formula_1>`为特殊处理范围指定公式。  
* `<Formula_2>`为非特殊处理范围指定公式。  



---
### English
# Excel-Formulas-by-Default

Fill the specified formula when the table cells in the specified range are null.  

Before using, please ensure that the spreadsheet file is in the format of a Microsoft 365MSO(versions 2504 Build 16.0.18730.20186) 64-bit Microsoft Excel Macro-Enabled Workbook (.xlsm).  
When you are using other softwares or file formats, or even different versions of Microsoft Office Excel, the VBA codes may not work.

<pre><code class="language-vba line-numbers">
' Worksheet change event handler: This function is triggered when a user modifies any table cells within the worksheet.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' Disable event handling to prevent the code from triggering itself and forming an infinite loop.
    Application.EnableEvents = False
    ' Check if the changed table cells is within the target range (Range).
    If Not Intersect(Target, Range("<Range_1>")) Is Nothing Then
        ' Loop through all modified table cells within the target range (Range_2 = Range_1).
        For Each cell In Intersect(Target, Range("<Range_2>"))
            ' If the table cells is cleared (value is null).
            If cell.Value = "" Then
                ' Declare a variable to store column letters (such as E or I).
                Dim colLetter As String
                ' Get the column letter part by splitting the table cells address (e.g., $E$4 is split into "E"). This line is required if the specified formula needs to reference the table cells address.
                colLetter = Split(cell.Address, "$")(1)
                ' Declare a variable to store the formula to be set.
                Dim formula As String
                ' Special handling for table cells in specified rows or columns (Row_value, Column_value).
                If cell.<Row,Column> = <Row_value,Column_value> Then
                    ' Set specified formulas for table cells in specified rows or columns.
                    formula = "=<Formula_1>"
                Else
                    ' Set specified formulas for table cells that are not in the specified rows or columns.If no additional processing of specified table cells is required, this layer of If_Else statement can be deleted.
                    formula = "=<Formula_2>"
                End If
                ' Set the constructed formula into the table cells.
                cell.formula = formula
            End If
        Next cell
    End If
    ' Restore event handling
    Application.EnableEvents = True
End Sub
</code></pre>

## `<>` are replacement tags, with detailed explanations as follows:

### `<Range_1>`&&`<Range_2>`
`<Range_1>` is the scope where the code takes effect; `<Range_2>` is the scope where the code executes. Generally, **these two values are equal**.  

#### Possible issues if they are not equal:

1. ##### Separation of event triggering conditions and processing scope 
    Triggering condition: The code only triggers when the user modifies cells within `<Range_1>` (judged by `Intersect(Target, Range("<Range_1>"))`).  
    Processing scope: However, the subsequent iteration is over cells within Range_2 (via `Intersect(Target, Range("<Range_2>"))`).  
    <br>
    Potential problems:  
    * If `<Range_2>` includes cells outside `<Range_1>`, modifications to these extra cells will not trigger the event, leaving them unprocessed by the code.  
    * If `<Range_2>` is a subset of `<Range_1>`, only modifications within the subset will be processed, while other parts are ignored.

2. ##### Abnormal formula application logic  
    Suppose the user modifies `<Range_1>` but not cells in `<Range_2>`:
    <br>
    * The code will iterate over the intersection of `<Range_2>` and `<Target>`, which may be empty in this case, causing the `<For Each>` loop to do nothing.  
    * Even if cells outside `<Range_2>` are cleared, the formula will not be applied.

3. ##### Example scenario
    Suppose：  
    * `Range_1 = "A1:B10"`(event triggering scope)  
    * `Range_2 = "C1:D10"` (formula application scope)  
    
    When the user modifies cells within `A1:B10`:  
    * The code triggers, but `Intersect(Target, Range("C1:D10"))` may be empty (since `Target` is within `A1:B10`), resulting in no formula being applied to any cells.

4. ##### Potential risks
    Infinite loop risk: If `<Range_2>` overlaps with `<Range_1>` and the formula calculation results may affect cells in `<Range_1>`, the event may be repeatedly triggered (even with `EnableEvents = False` protection).


### `<Row,Column>`
Select specified rows or columns.  
* If selecting rows, replace `<Row, Column> `directly with `Row`.
* If selecting columns, replace `<Row, Column>` directly with `Column`.

### `<Row_value,Column_value>`
The numeric value of the specified row or column; replace it directly with the corresponding number.  
If a column is specified, convert the **pure alphabetic base-27** notation to **base-10** notation before replacement.

### `<Formula_1>`&&`<Formula_2>`
* `<Formula_1>` is the formula specified for the special processing scope.
*`<Formula_2>` is the formula specified for the non-special processing scope.
---


### 日本語
# Excelのデフォルト値（数式で実現）

指定範囲内の表のセルが空の場合、指定された数式を自動的に入力します。

1. 使用する前に、スプレッドシートファイルがMicrosoft 365 MSO（バージョン2504 Build 16.0.18730.20186）64ビットのMicrosoft Excelのマクロを有効にしたブック形式（.xlsm）であることを確認してください。
2. 他のソフトウェアやファイル形式、さらには異なるバージョンのMicrosoft Office Excelでは、このVBAコードの有効性は保証されません。

```vba
' ワークシートの変更イベントハンドラ - ユーザーがワークシート内のセルを変更したときにトリガーされます。
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    ' イベント処理を無効にして、コードが自身をトリガーして無限ループになるのを防ぎます。
    Application.EnableEvents = False
    ' 変更されたセルが対象範囲（Range）内にあるかどうかを確認します。
    If Not Intersect(Target, Range("<Range_1>")) Is Nothing Then
        ' 対象範囲内で変更されたすべてのセル（Range_2 = Range_1）を繰り返し処理します。
        For Each cell In Intersect(Target, Range("<Range_2>"))
            ' セルが空になった場合（値が空文字列）。
            If cell.Value = "" Then
                ' 列のアルファベット（例：EまたはI）を格納する変数を宣言します。
                Dim colLetter As String
                ' セルのアドレスを分割して列のアルファベット部分を取得します（例：$E$4を "E" に分割）。指定された数式がセルアドレスを参照する場合はこの行が必要です。
                colLetter = Split(cell.Address, "$")(1)
                ' 設定する数式を格納する変数を宣言します。
                Dim formula As String
                ' 指定された行または列のセルを特殊処理します（Row_value,Column_value）。
                If cell.<Row,Column> = <Row_value,Column_value> Then
                    ' 指定された行または列のセルに指定された数式を設定します。
                    formula = "=<Formula_1>"
                Else
                    ' 指定された行または列以外のセルに指定された数式を設定します。指定されたセルに対する追加の処理が必要ない場合は、このIf_Else文を削除できます。
                    formula = "=<Formula_2>"
                End If
                ' 構築された数式をセルに設定します。
                cell.Formula = formula
            End If
        Next cell
    End If
    ' イベント処理を復元します。
    Application.EnableEvents = True
End Sub
```

## ここで`<>`は置き換えタグで、以下に詳細な説明を示します：

### `<Range_1>` && `<Range_2>`
`<Range_1>`はコードの有効範囲です；`<Range_2>`はコードの実行範囲で、通常はこの**2つの値は等しい**です。

#### もし等しくない場合、以下のような状況が発生する可能性があります：

1. ##### イベントトリガー条件と処理範囲の分離
    トリガー条件：コードはユーザーが`<Range_1>`内のセルを変更したときにのみトリガーされます（`Intersect(Target, Range("<Range_1>"))`で判断）。
    処理範囲：しかし、後続のループでは`<Range_2>`内のセルを走査します（`Intersect(Target, Range("<Range_2>"))`を通じて）。
    <br>
    考えられる問題：
    * `<Range_2>`に`<Range_1>`以外のセルが含まれている場合、これらの追加セルが変更されてもイベントはトリガーされず、コードが処理できなくなります。
    * `<Range_2>`が`<Range_1>`の部分集合である場合、部分集合内の変更のみが処理され、他の部分は無視されます。

2. ##### 数式適用ロジックの異常
    ユーザーが`<Range_1>`を変更したが、`<Range_2>`内のセルを変更しなかったと仮定します：
    * コードは`<Range_2>`と`<Target>`の共通部分を走査しますが、このとき共通部分が空の場合があり、`<For Each>`ループが何も実行しなくなります。
    * `<Range_2>`外のセルがクリアされても、数式は適用されません。

3. ##### サンプルシナリオ
    仮定：
    * `Range_1 = "A1:B10"`（イベントトリガー範囲）
    * `Range_2 = "C1:D10"`（数式適用範囲）

    ユーザーが`A1:B10`内のセルを変更したとき：
    * コードはトリガーされますが、`Intersect(Target, Range("C1:D10"))`が空の場合があり（`Target`は`A1:B10`内にあるため）、数式がどのセルにも適用されなくなります。

4. ##### 潜在的なリスク
    無限ループのリスク：`<Range_2>`と`<Range_1>`が重複し、数式の計算結果が`<Range_1>`内のセルに影響を与える可能性がある場合、イベントが繰り返しトリガーされる可能性があります（`EnableEvents = False`で保護しても）。

### `<Row,Column>`
指定された行または列を選択します。
* 行を選択する場合は、`<Row,Column>`を直接`Row`に置き換えます。
* 列を選択する場合は、`<Row,Column>`を直接`Column`に置き換えます。

### `<Row_value,Column_value>`
指定された行または列の数値を、それぞれの数字で直接置き換えます。
指定された列の場合、**純粋なアルファベットの27進数**を**10進数**に変換してから置き換えます。

### `<Formula_1>` && `<Formula_2>`
* `<Formula_1>`は特殊処理範囲に指定された数式です。
* `<Formula_2>`は非特殊処理範囲に指定された数式です。