---
title: ChartGroup.HasSeriesLines property (Excel)
keywords: vbaxl10.chm568082
f1_keywords:
- vbaxl10.chm568082
ms.prod: excel
api_name:
- Excel.ChartGroup.HasSeriesLines
ms.assetid: 4285cf5b-ebb0-a6fd-49c1-d36c341bd016
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.HasSeriesLines property (Excel)

**True** if a stacked column chart or bar chart has series lines, or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2D stacked bar, 2D stacked column, Pie of Pie, or Bar of Pie charts. Read/write **Boolean**.


## Syntax

_expression_.**HasSeriesLines**

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