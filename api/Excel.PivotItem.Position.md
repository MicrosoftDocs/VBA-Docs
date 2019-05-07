---
title: PivotItem.Position property (Excel)
keywords: vbaxl10.chm246081
f1_keywords:
- vbaxl10.chm246081
ms.prod: excel
api_name:
- Excel.PivotItem.Position
ms.assetid: 07e78622-f869-40d0-276a-b015ebe7a90f
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItem.Position property (Excel)

Returns or sets a **Long** value that represents the position of the item in its field, if the item is currently showing.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a **[PivotItem](Excel.PivotItem.md)** object.


## Example

This example displays the position number of the PivotTable item that contains the active cell.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The active item is in position number " & _ 
 ActiveCell.PivotItem.Position
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]