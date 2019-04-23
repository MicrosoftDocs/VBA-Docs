---
title: LegendKey property (Excel Graph)
keywords: vbagr10.chm65710
f1_keywords:
- vbagr10.chm65710
ms.prod: excel
api_name:
- Excel.LegendKey
ms.assetid: 55277508-2a81-c9c0-1f34-4d44c967ae8e
ms.date: 04/11/2019
localization_priority: Normal
---


# LegendKey property (Excel Graph)

Returns a **LegendKey** object that represents the legend key associated with the entry.

## Syntax

_expression_.**LegendKey**

_expression_ An expression that returns a **[LegendKey](Excel.LegendKey-graph-object.md)** object.


## Example

This example sets the legend key for legend entry one to be a triangle. The example should be run on a 2D line chart.


```vb
myChart.Legend.LegendEntries(1).LegendKey _ 
 .MarkerStyle = xlMarkerStyleTriangle
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]