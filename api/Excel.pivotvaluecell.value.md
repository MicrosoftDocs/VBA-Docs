---
title: PivotValueCell.Value property (Excel)
keywords: vbaxl10.chm918074
f1_keywords:
- vbaxl10.chm918074
ms.prod: excel
ms.assetid: 47bebd10-cd02-680f-f158-39c199e8ecf2
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotValueCell.Value property (Excel)

Returns the value at the location. The value is the value after **ShowAs** and other calculations have been applied. **Variant** can be **Empty**, **Number**, **Date**, **String**, or **Error** value.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[PivotValueCell](Excel.pivotvaluecell.md)** object.


## Remarks

This property works independently of whether the PivotTable is on a worksheet or not.


## Property value

**VARIANT**


## Example

This code sample uses the **PivotValueCell** property to test whether the value of one cell in a PivotTable is greater than another cell.

```vb
Sub TestEquality()
Dim X As Double
Dim Y As Double

'This code assumes that you have a Standalone PivotChart on one of the worksheets
X = ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).Value
Y = ThisWorkbook.PivotTables(1).PivotValueCell(1, 2).Value

If X > Y Then
MsgBox "X is greater than Y"
Else
MsgBox "Y is greater than X"
End If
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]