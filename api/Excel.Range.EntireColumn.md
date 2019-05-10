---
title: Range.EntireColumn property (Excel)
keywords: vbaxl10.chm144122
f1_keywords:
- vbaxl10.chm144122
ms.prod: excel
api_name:
- Excel.Range.EntireColumn
ms.assetid: 7be55670-75fd-fb02-dc1a-9d70e3a9d80d
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.EntireColumn property (Excel)

Returns a **Range** object that represents the entire column (or columns) that contains the specified range. Read-only.


## Syntax

_expression_.**EntireColumn**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example sets the value of the first cell in the column that contains the active cell. The example must be run from a worksheet.

```vb
ActiveCell.EntireColumn.Cells(1, 1).Value = 5
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
