---
title: ChartGroup.RadarAxisLabels property (Excel)
keywords: vbaxl10.chm568087
f1_keywords:
- vbaxl10.chm568087
ms.prod: excel
api_name:
- Excel.ChartGroup.RadarAxisLabels
ms.assetid: 36bb1e30-99b0-e795-2730-145421a2a342
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.RadarAxisLabels property (Excel)

Returns a **[TickLabels](Excel.TickLabels(object).md)** object that represents the radar axis labels for the specified chart group. Read-only.


## Syntax

_expression_.**RadarAxisLabels**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example turns on radar axis labels for chart group one on Chart1, and then sets the color for the labels. The example should be run on a radar chart.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasRadarAxisLabels = True 
 .RadarAxisLabels.Font.ColorIndex = 3 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]