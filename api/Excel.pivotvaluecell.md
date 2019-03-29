---
title: PivotValueCell object (Excel)
keywords: vbaxl10.chm917072
f1_keywords:
- vbaxl10.chm917072
ms.prod: excel
ms.assetid: 1857160d-9eab-d026-ef7d-af6187c6490e
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotValueCell object (Excel)

Provides a way to expose values of cells in the case that actual cells (**Range** objects) are not available.


## Example

This code sample uses the **PivotValueCell** property to test whether the value of one cell in a PivotTable is greater than another cell.

```vb
Sub TestEquality()
Dim X As Double
Dim Y As Double

'This code assumes that you have a Standalone PivotChart on one of the worksheets.
X = ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).Value
Y = ThisWorkbook.PivotTables(1).PivotValueCell(1, 2).Value

If X > Y Then
MsgBox "X is greater than Y"
Else
MsgBox "Y is greater than X"
End If
End Sub
```

## Methods

- [ShowDetail](Excel.pivotvaluecell.showdetail.md)

## Properties

- [Application](Excel.pivotvaluecell.application.md)
- [Creator](Excel.pivotvaluecell.creator.md)
- [Parent](Excel.pivotvaluecell.parent.md)
- [PivotCell](Excel.pivotvaluecell.pivotcell.md)
- [ServerActions](Excel.pivotvaluecell.serveractions.md)
- [Value](Excel.pivotvaluecell.value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]