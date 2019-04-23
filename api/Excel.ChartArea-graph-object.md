---
title: ChartArea object (Excel Graph)
keywords: vbagr10.chm5207179
f1_keywords:
- vbagr10.chm5207179
ms.prod: excel
api_name:
- Excel.ChartArea
ms.assetid: 85fcf460-6b2b-142f-ce4a-4a74e9d8efd3
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartArea object (Excel Graph)

Represents the chart area of the specified chart. 

## Remarks

The chart area in a 2D chart contains the axes, the chart title, the axis titles, and the legend. 

The chart area in a 3D chart contains the chart title and the legend; it doesn't include the plot area (the area within the chart area where the data is plotted). 

For information about formatting the plot area, see the **[PlotArea](Excel.PlotArea-graph-object.md)** object.

Use the **[ChartArea](excel.chartarea-graph-property.md)** property to return the **ChartArea** object. 

## Example

The following example sets the pattern for the chart area.

```vb
myChart.ChartArea.Interior.Pattern = xlLightDown
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]