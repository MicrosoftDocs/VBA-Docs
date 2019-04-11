---
title: Legend property (Excel Graph)
keywords: vbagr10.chm5207602
f1_keywords:
- vbagr10.chm5207602
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: 03d13546-c567-04b3-8ed5-cb99dc97c8e4
ms.date: 04/11/2019
localization_priority: Normal
---


# Legend property (Excel Graph)

Returns a **Legend** object that represents the legend for the specified chart. Read-only.

## Syntax

_expression_.**Legend**

_expression_ An expression that returns a **[Legend](Excel.Legend-graph-object.md)** object.

## Example

This example turns on the legend for the chart, and then sets the font color for the legend to blue.

```vb
myChart.HasLegend = True 
myChart.Legend.Font.ColorIndex = 5
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]