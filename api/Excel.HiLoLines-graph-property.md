---
title: HiLoLines property (Excel Graph)
keywords: vbagr10.chm65679
f1_keywords:
- vbagr10.chm65679
ms.prod: excel
api_name:
- Excel.HiLoLines
ms.assetid: ed2ff722-b477-4346-d807-3d2615abd845
ms.date: 04/11/2019
localization_priority: Normal
---


# HiLoLines property (Excel Graph)

Returns a **HiLoLines** object that represents the high-low lines for the specified series on a line chart. Applies only to line charts. Read-only.

## Syntax

_expression_.**HiLoLines**

_expression_ An expression that returns a **[HiLoLines](Excel.HiLoLines-graph-object.md)** object.

## Example

This example turns on high-low lines for chart group one on the chart and then sets their line style, weight, and color. The example should be run on a 2D line chart that has three series of stock-quote-like data (high-low-close).

```vb
With myChart.ChartGroups(1) 
 .HasHiLoLines = True 
 With .HiLoLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]