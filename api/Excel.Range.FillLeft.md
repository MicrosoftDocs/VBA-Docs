---
title: Range.FillLeft method (Excel)
keywords: vbaxl10.chm144125
f1_keywords:
- vbaxl10.chm144125
ms.prod: excel
api_name:
- Excel.Range.FillLeft
ms.assetid: 42722b18-8b40-c27b-8bca-ef180cf0f636
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.FillLeft method (Excel)

Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.


## Syntax

_expression_.**FillLeft**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

Variant


## Example

This example fills the range A1:M1 on Sheet1, based on the contents of cell M1.

```vb
Worksheets("Sheet1").Range("A1:M1").FillLeft
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]