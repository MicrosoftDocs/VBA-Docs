---
title: Range.Text property (Excel)
keywords: vbaxl10.chm144209
f1_keywords:
- vbaxl10.chm144209
ms.prod: excel
api_name:
- Excel.Range.Text
ms.assetid: e38c15b1-5941-0a28-1acf-328bc214a2e0
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Text property (Excel)

Returns the formatted text for the specified object. Read-only **String**.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The **Text** property is most often used for a range of one cell. If the range includes more than one cell, the **Text** property returns **Null**, except when all the cells in the range have identical contents and formats.

If the contents of the cell is wider than the width available for display, the **Text** property will modify the displayed value.

## Property Differences Example

This example illustrates the difference between the **Text** and **[Value](Excel.Range.Value.md)**  properties of cells that contain formatted numbers.

```vb
Option Explicit

Public Sub DifferencesBetweenValueAndTextProperties()
    Dim cell As Range
    Set cell = Worksheets("Sheet1").Range("A1")
    cell.Value = 1198.3
    cell.NumberFormat = "$#,##0_);($#,##0)"
    
    MsgBox "'" & cell.Value & "' is the value." 'Returns: "'1198.3' is the value."
    MsgBox "'" & cell.Text & "' is the text."    'Returns: "'$1,198' is the text."
End Sub
```

## Text Width Differences

Cells containing numeric values may have their displayed value modified when the column isn't wide enough. The example below shows this using two columns. The first column is wide enough to display the values. A format is applied and then a value entered showing the full value. The second column has its width reduced such that when the cells are copied over it is too narrow causing the displayed value to be moified.

```vb
Public Sub TextWidthDifferences()
    
    Dim wideColumn As Range
    Set wideColumn = Sheet1.Range("B2")
    wideColumn.Value = "Wide Enough Column"
    wideColumn.Columns.AutoFit
    
    Sheet1.Range("B3").Value2 = 123456789
    
    Const CurrencyWith2DecimalsFormat As String = "$#,##0.00"
    Dim currencyCell As Range
    Set currencyCell = Sheet1.Range("B4")
    currencyCell.Value2 = 1234.56
    currencyCell.NumberFormat = CurrencyWith2DecimalsFormat
    
    Dim narrowColumn As Range
    Set narrowColumn = Sheet1.Range("C2")
    narrowColumn.Value = "Reduced Width Column"
    narrowColumn.ColumnWidth = 7.5
    
    Sheet1.Range("B3:B4").AutoFill Destination:=Sheet1.Range("B3:C4"), Type:=XlAutoFillType.xlFillDefault
    Debug.Print Sheet1.Range("C3").Text
    Debug.Print Sheet1.Range("C4").Text
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
