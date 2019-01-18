---
title: SeriesLines object (Excel)
keywords: vbaxl10.chm597072
f1_keywords:
- vbaxl10.chm597072
ms.prod: excel
api_name:
- Excel.SeriesLines
ms.assetid: db044358-d14b-ef45-4e42-237b8ee46ff0
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesLines object (Excel)

Represents series lines in a chart group.


## Remarks

 Series lines connect the data values from each series. Only 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts can have series lines. This object isn't a collection. There's no object that represents a single series line; you either have series lines turned on for all points in a chart group or you have them turned off.

If the  **[HasSeriesLines](Excel.ChartGroup.HasSeriesLines.md)** property is **False** , most properties of the **SeriesLines** object are disabled.


## Example

Use the  **SeriesLines** property to return a **SeriesLines** object. The following example adds series lines to chart group one in embedded chart one on worksheet one (the chart must be a 2-D stacked bar or column chart).


```vb
With Worksheets(1).ChartObjects(1).Chart.ChartGroups(1) 
 .HasSeriesLines = True 
 .SeriesLines.Border.Color = RGB(0, 0, 255) 
End With
```


## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)


