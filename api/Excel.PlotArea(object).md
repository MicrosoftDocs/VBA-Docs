---
title: PlotArea object (Excel)
keywords: vbaxl10.chm617072
f1_keywords:
- vbaxl10.chm617072
ms.prod: excel
api_name:
- Excel.PlotArea
ms.assetid: 85c42124-268c-8b0e-ba5d-c2f6fbf53e79
ms.date: 03/30/2019
localization_priority: Normal
---


# PlotArea object (Excel)

Represents the plot area of a chart.


## Remarks

This is the area where your chart data is plotted. The plot area on a 2D chart contains the data markers, gridlines, data labels, trendlines, and optional chart items placed in the chart area. The plot area on a 3D chart contains all the above items plus the walls, floor, axes, axis titles, and tick-mark labels in the chart.

The plot area is surrounded by the chart area. The chart area on a 2D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3D chart contains the chart title and the legend. For information about formatting the chart area, see the **[ChartArea](Excel.ChartArea(object).md)** object.


## Example

Use the **PlotArea** property to return a **PlotArea** object. The following example activates the chart sheet named "Chart1," places a dashed border around the chart area of the active chart, and places a dotted border around the plot area.

```vb
Charts("Chart1").Activate 
With ActiveChart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
End With
```


## Methods

- [ClearFormats](Excel.PlotArea.ClearFormats.md)
- [Select](Excel.PlotArea.Select.md)

## Properties

- [Application](Excel.PlotArea.Application.md)
- [Creator](Excel.PlotArea.Creator.md)
- [Format](Excel.PlotArea.Format.md)
- [Height](Excel.PlotArea.Height.md)
- [InsideHeight](Excel.PlotArea.InsideHeight.md)
- [InsideLeft](Excel.PlotArea.InsideLeft.md)
- [InsideTop](Excel.PlotArea.InsideTop.md)
- [InsideWidth](Excel.PlotArea.InsideWidth.md)
- [Left](Excel.PlotArea.Left.md)
- [Name](Excel.PlotArea.Name.md)
- [Parent](Excel.PlotArea.Parent.md)
- [Position](Excel.PlotArea.Position.md)
- [Top](Excel.PlotArea.Top.md)
- [Width](Excel.PlotArea.Width.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]