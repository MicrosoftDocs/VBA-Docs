---
title: ChartGroup.HasRadarAxisLabels property (Excel)
keywords: vbaxl10.chm568081
f1_keywords:
- vbaxl10.chm568081
ms.prod: excel
api_name:
- Excel.ChartGroup.HasRadarAxisLabels
ms.assetid: 7b3e0a6f-00da-ac8b-9a64-d79923f13481
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.HasRadarAxisLabels property (Excel)

**True** if a radar chart has axis labels. Applies only to radar charts. Read/write **Boolean**.


## Syntax

_expression_.**HasRadarAxisLabels**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example turns on radar axis labels for chart group one on Chart1 and sets their color. The example should be run on a radar chart.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]