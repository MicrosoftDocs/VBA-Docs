---
title: ChartGroup.SeriesLines property (Word)
keywords: vbawd10.chm263454746
f1_keywords:
- vbawd10.chm263454746
ms.prod: word
api_name:
- Word.ChartGroup.SeriesLines
ms.assetid: 23f36b19-99ed-f4d5-23b5-a8cd35bbf75c
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SeriesLines property (Word)

Returns the series lines for a 2D stacked bar, 2D stacked column, pie-of-pie, or bar-of-pie chart. Read-only  **[SeriesLines](Word.SeriesLines.md)**.


## Syntax

_expression_.**SeriesLines**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Example

The following example enables series lines for chart group one of the first chart in the active document, and then sets the line style, weight, and color of the series lines. You should run the example on a 2D stacked column chart that has two or more series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasSeriesLines = True 
 With .SeriesLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
 End With 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]