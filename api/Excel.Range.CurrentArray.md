---
title: Range.CurrentArray property (Excel)
keywords: vbaxl10.chm144110
f1_keywords:
- vbaxl10.chm144110
api_name:
- Excel.Range.CurrentArray
ms.assetid: 147f8834-5aef-900f-75de-df91a6a76005
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.CurrentArray property (Excel)

If the specified cell is part of an array, returns a **Range** object that represents the entire array. Read-only.


## Syntax

_expression_.**CurrentArray**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example assumes that cell A1 on Sheet1 is the active cell, and that the active cell is part of an array that includes cells A1:A10. The example selects cells A1:A10 on Sheet1.

```vb
ActiveCell.CurrentArray.Select
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
