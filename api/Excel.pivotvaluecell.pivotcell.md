---
title: PivotValueCell.PivotCell property (Excel)
keywords: vbaxl10.chm918073
f1_keywords:
- vbaxl10.chm918073
ms.prod: excel
ms.assetid: 18fa81bd-3169-9f08-9418-93ea5443efb2
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotValueCell.PivotCell property (Excel)

Returns the [PivotCell object (Excel)](Excel.PivotCell.md) that specifies the location of the **PivotValueCell**. Read-only


## Syntax

_expression_. `PivotCell`

_expression_ A variable that represents a [PivotValueCell object (Excel)](Excel.pivotvaluecell.md) object.


## Example

The following code sample uses the  **PivotCell** property to get the Multi-dimensional Expressions (MDX) expression for the specified cell.


```vb
Sub GetMDX()
   'Get the MDX query for a particular PivotCell in a workbook level PivotTable
   MsgBox "The MDX for the PivotCell 1, 1: " & _
   ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).PivotCell.MDX
End Sub
```


## Property value

 **PIVOTCELL**


## See also



[PivotValueCell Object](Excel.pivotvaluecell.md)

