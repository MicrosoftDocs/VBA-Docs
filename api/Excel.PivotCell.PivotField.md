---
title: PivotCell.PivotField property (Excel)
keywords: vbaxl10.chm692076
f1_keywords:
- vbaxl10.chm692076
api_name:
- Excel.PivotCell.PivotField
ms.assetid: a1217848-e3b0-0e92-168b-3a9c21245380
ms.date: 05/04/2019
ms.localizationpriority: medium
---


# PivotCell.PivotField property (Excel)

Returns a **[PivotField](Excel.PivotField.md)** object that represents the PivotTable field containing the upper-left corner of the specified range.


## Syntax

_expression_.**PivotField**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Example

This example displays the name of the PivotTable field that contains the active cell.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the field " & _ 
 ActiveCell.PivotField.Name
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]