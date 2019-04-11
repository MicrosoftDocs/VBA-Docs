---
title: SeriesLines object (Word)
keywords: vbawd10.chm3093
f1_keywords:
- vbawd10.chm3093
ms.prod: word
api_name:
- Word.SeriesLines
ms.assetid: 7521c592-c5aa-8e50-6268-840a41b3a282
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesLines object (Word)

Represents series lines in a chart group.


## Remarks

 Series lines connect the data values from each series. Only 2D stacked bar, 2D stacked column, pie-of-pie, or bar-of-pie charts can have series lines. This object is not a collection. There is no object that represents a single series line; you either enable series lines for all points in a chart group or you disable them.

If the  **[HasSeriesLines](Word.ChartGroup.HasSeriesLines.md)** property is **False**, most properties of the **SeriesLines** object are disabled.


## Example

Use the  **[SeriesLines](Word.ChartGroup.SeriesLines.md)** property to return a **SeriesLines** object. The following example adds series lines to chart group one in embedded chart one on worksheet one (the chart must be a 2D stacked bar or column chart).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasSeriesLines = True 
 .SeriesLines.Border.Color = RGB(0, 0, 255) 
 End With 
 End If 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]