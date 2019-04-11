---
title: PlotArea object (Word)
keywords: vbawd10.chm816
f1_keywords:
- vbawd10.chm816
ms.prod: word
api_name:
- Word.PlotArea
ms.assetid: 72d30767-7cfc-3063-0b49-f9fbc129a52c
ms.date: 06/08/2017
localization_priority: Normal
---


# PlotArea object (Word)

Represents the plot area of a chart.


## Remarks

 This is the area where your chart data is plotted. The plot area on a 2D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.

The plot area is surrounded by the chart area. The chart area on a 2D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3D chart contains the chart title and the legend. For information about formatting the chart area, see the  **[ChartArea](Word.ChartArea.md)** object.


## Example

Use the  **[PlotArea](Word.Chart.PlotArea.md)** property to return a **PlotArea** object. The following example places a dashed border around the chart area of the first chart in the active document, and then places a dotted border around the plot area.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
 End With 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]