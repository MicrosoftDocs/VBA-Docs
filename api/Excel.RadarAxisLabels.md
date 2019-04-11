---
title: RadarAxisLabels property (Excel Graph)
keywords: vbagr10.chm65680
f1_keywords:
- vbagr10.chm65680
ms.prod: excel
api_name:
- Excel.RadarAxisLabels
ms.assetid: e382e92c-96f2-a9ee-720f-dcb85e5e2e7c
ms.date: 04/12/2019
localization_priority: Normal
---


# RadarAxisLabels property (Excel Graph)

Returns a **TickLabels** object that represents the radar axis labels for the specified chart group. Read-only.

## Syntax

_expression_.**RadarAxisLabels**

_expression_ An expression that returns a **[TickLabels](Excel.TickLabels-graph-object.md)** object.

## Example

This example turns on radar axis labels for chart group one on the chart, and then sets the color for the labels. The example should be run on a radar chart.

```vb
With myChart.ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]