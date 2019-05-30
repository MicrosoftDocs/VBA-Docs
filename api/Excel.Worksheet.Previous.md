---
title: Worksheet.Previous property (Excel)
keywords: vbaxl10.chm174086
f1_keywords:
- vbaxl10.chm174086
ms.prod: excel
api_name:
- Excel.Worksheet.Previous
ms.assetid: 8409e3c6-564e-2ba1-1e49-79a1c37cc845
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Previous property (Excel)

Returns a **Worksheet** object that represents the previous sheet.


## Syntax

_expression_.**Previous**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


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