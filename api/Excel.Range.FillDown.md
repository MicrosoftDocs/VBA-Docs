---
title: Range.FillDown method (Excel)
keywords: vbaxl10.chm144124
f1_keywords:
- vbaxl10.chm144124
api_name:
- Excel.Range.FillDown
ms.assetid: bb7c0b2d-8dd9-13e5-b90a-b2708935afa9
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.FillDown method (Excel)

Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.


## Syntax

_expression_.**FillDown**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

Variant


## Example

This example fills the range A1:A10 on Sheet1, based on the contents of cell A1.

```vb
Worksheets("Sheet1").Range("A1:A10").FillDown
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
