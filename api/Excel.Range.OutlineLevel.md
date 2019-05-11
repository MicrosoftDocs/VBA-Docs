---
title: Range.OutlineLevel property (Excel)
keywords: vbaxl10.chm144171
f1_keywords:
- vbaxl10.chm144171
ms.prod: excel
api_name:
- Excel.Range.OutlineLevel
ms.assetid: bdab08a4-3576-4a65-2556-43ed9e9a576e
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.OutlineLevel property (Excel)

Returns or sets the current outline level of the specified row or column. Read/write **Variant**.


## Syntax

_expression_.**OutlineLevel**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

Level one is the outermost summary level.


## Example

This example sets the outline level for row two on Sheet1.

```vb
Worksheets("Sheet1").Rows(2).OutlineLevel = 1
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
