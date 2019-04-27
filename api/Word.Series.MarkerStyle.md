---
title: Series.MarkerStyle property (Word)
keywords: vbawd10.chm123732040
f1_keywords:
- vbawd10.chm123732040
ms.prod: word
api_name:
- Word.Series.MarkerStyle
ms.assetid: d9ba7847-2785-0f29-7e6e-d4bb2d62fc2f
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.MarkerStyle property (Word)

Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write **[XlMarkerStyle](Word.xlmarkerstyle.md)**.


## Syntax

_expression_.**MarkerStyle**

_expression_ A variable that represents a '[Series](Word.Series.md)' object.



## Example

The following example sets the marker style for series one for the first chart in the active document. You should run the example on a 2D line chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle 
 End If 
End With
```


## See also


[Series Object](Word.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]