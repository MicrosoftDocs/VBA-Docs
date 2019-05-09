---
title: PivotTable.PivotValueCell method (Excel)
keywords: vbaxl10.chm235203
f1_keywords:
- vbaxl10.chm235203
ms.prod: excel
ms.assetid: 9edb96f1-f728-de21-bcc2-e8f0e9110b74
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PivotValueCell method (Excel)

Retrieve the **[PivotValueCell](Excel.pivotvaluecell.md)** object for a given PivotTable provided certain row and column indices.


## Syntax

_expression_.**PivotValueCell** (_RowLine_, _ColumnLine_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowLine_|Optional|**Variant**|If of type **PivotLine**, specifies the **PivotLine** in the row area that the **PivotValueCell** is aligned with. If of type **Int**, specifies the position of the **PivotLine** on the row area that the **PivotValueCell** is aligned with. If missing, Empty, Null, or 0 specifies the grand total row.|
| _ColumnLine_|Optional|**Variant**|If of type **PivotLine**, specifies the **PivotLine** in the column area that the **PivotValueCell** is aligned with. If of type **Int**, specifies the position of the **PivotLine** on the column area that the **PivotValueCell** is aligned with. If missing, Empty, Null or 0 specifies the grand total column.|

## Return value

**PIVOTVALUECELL**


## Example

This code sample uses the **PivotValueCell** property to test whether the value of one cell in a PivotTable is greater than the value of another cell.

```vb
Sub TestEquality()
Dim X As Double
Dim Y As Double

'This code assumes you have a Standalone PivotChart on one of the worksheets
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