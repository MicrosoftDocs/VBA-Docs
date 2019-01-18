---
title: Points Object (PowerPoint)
keywords: vbapp10.chm715000
f1_keywords:
- vbapp10.chm715000
ms.prod: powerpoint
api_name:
- PowerPoint.Points
ms.assetid: f3ee69d3-ab8f-e300-bbf4-00ea97d47c2a
ms.date: 06/08/2017
localization_priority: Normal
---


# Points Object (PowerPoint)

A collection of all the  **[Point](PowerPoint.Point.md)** objects in the specified series in a chart.


## Remarks

Use  **[Points](PowerPoint.Series.Points.md)** ( _Index_ ), where _Index_ is the point index number, to return a single **Point** object. Points are numbered from left to right on the series. `Points(1)` is the leftmost point, and `Points(Points.Count)` is the rightmost point.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **Points** method to return the **Points** collection. The following example adds a data label to the last point in series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Points

            .Item(.Count).ApplyDataLabels Type:=xlShowValue

        End With

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 The following example sets the marker style for the third point in series one for the first chart in the active document. The specified series must be a 2-D line, scatter, or radar series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

