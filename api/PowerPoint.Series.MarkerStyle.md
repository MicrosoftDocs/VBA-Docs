---
title: Series.MarkerStyle property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.MarkerStyle
ms.assetid: e985978e-f0cf-b809-ebe1-f5504e9e8df6
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.MarkerStyle property (PowerPoint)

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write **[XlMarkerStyle](PowerPoint.XlMarkerStyle.md)**.


## Syntax

_expression_.**MarkerStyle**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Example

> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the marker style for series one for the first chart in the active document. You should run the example on a 2D line chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle

    End If

End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]