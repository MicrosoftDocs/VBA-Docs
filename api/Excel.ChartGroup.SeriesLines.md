---
title: ChartGroup.SeriesLines property (Excel)
keywords: vbaxl10.chm568089
f1_keywords:
- vbaxl10.chm568089
ms.prod: excel
api_name:
- Excel.ChartGroup.SeriesLines
ms.assetid: 3e2156c3-c4dd-ef22-1645-ba27e7b499b8
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.SeriesLines property (Excel)

Returns a **[SeriesLines](Excel.SeriesLines(object).md)** object that represents the series lines for a 2D stacked bar, 2D stacked column, Pie of Pie, or Bar of Pie chart. Read-only.


## Syntax

_expression_.**SeriesLines**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example turns on series lines for chart group one on Chart1, and then sets their line style, weight, and color. The example should be run on a 2D stacked column chart that has two or more series.

```vb
With Charts("Chart1").ChartGroups(1) 
 .HasSeriesLines = True 
 With .SeriesLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]