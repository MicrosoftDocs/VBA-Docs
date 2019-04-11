---
title: SeriesLines object (Excel Graph)
keywords: vbagr10.chm131201
f1_keywords:
- vbagr10.chm131201
ms.prod: excel
api_name:
- Excel.SeriesLines
ms.assetid: 958145eb-8801-b285-b3b4-99fd7b7882ed
ms.date: 04/06/2019
localization_priority: Normal
---


# SeriesLines object (Excel Graph)

Represents series lines in the specified chart group. Series lines connect the data values in each series. Only 2D stacked-bar or column chart groups can have series lines. 

This object isn't a collection. There's no object that represents a single series line; either you have series lines turned on for all points in a chart group or you have them turned off.  

## Remarks

Use the **[SeriesLines](excel.serieslines-graph-property.md)** property to return the **SeriesLines** object. 

If the **[HasSeriesLines](Excel.HasSeriesLines.md)** property is **False**, most properties of the **SeriesLines** object are disabled.

## Example

The following example adds series lines to chart group one in the chart. The chart must be a 2D stacked-bar or column chart.

```vb
With myChart.ChartGroups(1) 
 .HasSeriesLines = True 
 .SeriesLines.Border.Color = RGB(0, 0, 255) 
End With
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]