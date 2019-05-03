---
title: PivotCell.PivotItem property (Excel)
keywords: vbaxl10.chm692077
f1_keywords:
- vbaxl10.chm692077
ms.prod: excel
api_name:
- Excel.PivotCell.PivotItem
ms.assetid: 3b131e96-8589-9d72-d4d9-afe2d3d6137c
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.PivotItem property (Excel)

Returns a **[PivotItem](Excel.PivotItem.md)** object that represents the PivotTable item containing the upper-left corner of the specified range.


## Syntax

_expression_.**PivotItem**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Example

This example displays the name of the PivotTable item that contains the active cell on Sheet1.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the item " & _ 
 ActiveCell.PivotItem.Name
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]