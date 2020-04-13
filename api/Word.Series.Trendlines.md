---
title: Series.Trendlines method (Word)
keywords: vbawd10.chm123732122
f1_keywords:
- vbawd10.chm123732122
ms.prod: word
api_name:
- Word.Series.Trendlines
ms.assetid: 300dca01-097f-8a3d-4f63-a1841a92098e
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.Trendlines method (Word)

Returns a collection of all the trendlines for the series.


## Syntax

_expression_.**Trendlines** (_Index_)

_expression_ A variable that represents a '[Series](Word.Series.md)' object.


## Return value

A **[Trendlines](Word.Trendlines.md)** object that represents all the treadlines for the series.


## Example

The following example adds a linear trendline to series one for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines.Add Type:=xlLinear 
 End If 
End With
```


## See also


[Series Object](Word.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]