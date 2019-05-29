---
title: Range.Previous property (Excel)
keywords: vbaxl10.chm144180
f1_keywords:
- vbaxl10.chm144180
ms.prod: excel
api_name:
- Excel.Range.Previous
ms.assetid: 6ee986eb-9242-63f3-6885-1ad3730f106b
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Previous property (Excel)

Returns a **Range** object that represents the previous cell.


## Syntax

_expression_.**Previous**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

If the object is a range, this property emulates pressing Shift+Tab; unlike the key combination, however, the property returns the previous cell without selecting it.

On a protected sheet, this property returns the previous unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the left of the specified cell.


## Example

This example selects the previous unlocked cell on Sheet1. If Sheet1 is unprotected, this is the cell immediately to the left of the active cell.

```vb
Worksheets("Sheet1").Activate 
ActiveCell.Previous.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]