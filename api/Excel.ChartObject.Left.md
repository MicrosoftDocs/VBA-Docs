---
title: ChartObject.Left property (Excel)
keywords: vbaxl10.chm494084
f1_keywords:
- vbaxl10.chm494084
ms.prod: excel
api_name:
- Excel.ChartObject.Left
ms.assetid: 2b4964e2-624e-e53e-6efc-f792bf28a202
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.Left property (Excel)

Returns or sets a **Double** value that represents the distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).


## Syntax

_expression_.**Left**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example aligns the left edge of the embedded chart with the left edge of column B.

```vb
With Worksheets("Sheet1") 
 .ChartObjects(1).Left = .Columns("B").Left 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]