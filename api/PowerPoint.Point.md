---
title: Point object (PowerPoint)
keywords: vbapp10.chm714000
f1_keywords:
- vbapp10.chm714000
ms.prod: powerpoint
api_name:
- PowerPoint.Point
ms.assetid: e0137fdd-5632-88d7-a6c0-57a76717e736
ms.date: 06/08/2017
localization_priority: Normal
---


# Point object (PowerPoint)

Represents a single point in a series in a chart.


## Remarks

 The **Point** object is a member of the **[Points](PowerPoint.Points.md)** collection. The **Points** collection contains all the points in one series.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[Points](PowerPoint.Series.Points.md)** (_index_), where _index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point. The following example sets the marker style for the third point in series one for the first chart in the active document. The specified series must be a 2D line, scatter, or radar series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond

    End If

End With


```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]