---
title: ChartObject.TopLeftCell property (Excel)
keywords: vbaxl10.chm494093
f1_keywords:
- vbaxl10.chm494093
ms.prod: excel
api_name:
- Excel.ChartObject.TopLeftCell
ms.assetid: 582879c6-528d-3979-c52e-13c738ba6902
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.TopLeftCell property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the cell that lies under the upper-left corner of the specified object. Read-only.


## Syntax

_expression_.**TopLeftCell**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example displays the address of the cell beneath the upper-left corner of embedded chart one on Sheet1.

```vb
MsgBox "The top-left corner is over cell " & _ 
 Worksheets("Sheet1").ChartObjects(1).TopLeftCell.Address
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]