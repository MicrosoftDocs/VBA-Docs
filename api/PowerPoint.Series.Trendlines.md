---
title: Series.Trendlines method (PowerPoint)
keywords: vbapp10.chm65690
f1_keywords:
- vbapp10.chm65690
ms.prod: powerpoint
api_name:
- PowerPoint.Series.Trendlines
ms.assetid: 17578607-d0aa-dcc2-1eec-3af031f17c2d
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.Trendlines method (PowerPoint)

Returns a collection of all the trendlines for the series.


## Syntax

_expression_.**Trendlines** (_Index_)

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Return value

A **[Trendlines](PowerPoint.Trendlines.md)** object that represents all the treadlines for the series.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a linear trendline to series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Trendlines.Add Type:=xlLinear

    End If

End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]