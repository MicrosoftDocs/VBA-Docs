---
title: Range.PivotItem property (Excel)
keywords: vbaxl10.chm144176
f1_keywords:
- vbaxl10.chm144176
ms.prod: excel
api_name:
- Excel.Range.PivotItem
ms.assetid: 02a41786-074b-ae34-5d2c-407006fe526d
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.PivotItem property (Excel)

Returns a **[PivotItem](Excel.PivotItem.md)** object that represents the PivotTable item containing the upper-left corner of the specified range.


## Syntax

_expression_.**PivotItem**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example displays the name of the PivotTable item that contains the active cell on Sheet1.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the item " & _ 
 ActiveCell.PivotItem.Name
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]