---
title: NumberFormat property (Excel Graph)
keywords: vbagr10.chm65729
f1_keywords:
- vbagr10.chm65729
ms.prod: excel
api_name:
- Excel.NumberFormat
ms.assetid: 0a8b652a-6c8d-d4bd-4e93-e62ca86e6053
ms.date: 04/12/2019
localization_priority: Normal
---


# NumberFormat property (Excel Graph)

Returns or sets the format code for the object. Returns **Null** if the cells in the specified range don't all have the same number format. Read/write **String** for all objects, except for the **[Range](excel.range-graph-object.md)** object, which is read/write **Variant**.

## Syntax

_expression_.**NumberFormat**

_expression_ Required. An expression that returns one of the above objects.


## Example

This example sets the number format for the data labels for series one.

```vb
myChart.SeriesCollection(1).DataLabels.NumberFormat = "General"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]