---
title: Range.HasArray property (Excel)
keywords: vbaxl10.chm144142
f1_keywords:
- vbaxl10.chm144142
ms.prod: excel
api_name:
- Excel.Range.HasArray
ms.assetid: fac17206-8671-6209-9133-d56da6ea2b9c
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.HasArray property (Excel)

**True** if the specified cell is part of an array formula. Read-only **Variant**.


## Syntax

_expression_.**HasArray**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example displays a message if the active cell on Sheet1 is part of an array.

```vb
Worksheets("Sheet1").Activate 
If ActiveCell.HasArray =True Then 
 MsgBox "The active cell is part of an array" 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]