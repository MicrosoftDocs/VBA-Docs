---
title: HasRadarAxisLabels property (Excel Graph)
keywords: vbagr10.chm65600
f1_keywords:
- vbagr10.chm65600
ms.prod: excel
api_name:
- Excel.HasRadarAxisLabels
ms.assetid: 8baa636a-262c-15b4-f8d5-94d77a8101c5
ms.date: 04/11/2019
localization_priority: Normal
---


# HasRadarAxisLabels property (Excel Graph)

**True** if a radar chart has axis labels. Applies only to radar charts. Read/write **Boolean**.

## Syntax

_expression_.**HasRadarAxisLabels**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example turns on radar axis labels for chart group one and sets their color. The example should be run on a radar chart.

```vb
With myChart.ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]