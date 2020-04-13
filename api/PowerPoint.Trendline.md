---
title: Trendline object (PowerPoint)
keywords: vbapp10.chm720000
f1_keywords:
- vbapp10.chm720000
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline
ms.assetid: 74755c19-0a7d-cbbf-857e-78740adf6aa4
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline object (PowerPoint)

Represents a trendline in a chart.


## Remarks

A trendline shows the trend, or direction, of data in a series. The **Trendline** object is a member of the **[Trendlines](PowerPoint.Trendlines.md)** collection. The **Trendlines** collection contains all the **Trendline** objects for a single series.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[Trendlines](PowerPoint.Series.Trendlines.md)** (Index), where Index is the trendline index number, to return a single **Trendline** object.

The index number denotes the order in which the trendlines were added to the series.  `Trendlines(1)` is the first trendline added to the series, and `Trendlines(Trendlines.Count)` is the last one added.

The following example changes the trendline type for the first series of the first chart in the active document. If the series has no trendline, this example will fail.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]