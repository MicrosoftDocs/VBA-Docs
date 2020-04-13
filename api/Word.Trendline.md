---
title: Trendline object (Word)
keywords: vbawd10.chm402
f1_keywords:
- vbawd10.chm402
ms.prod: word
api_name:
- Word.Trendline
ms.assetid: 1cfe897f-26ad-a838-ed9b-f3fd945ff7ea
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline object (Word)

Represents a trendline in a chart.


## Remarks

A trendline shows the trend, or direction, of data in a series. The **Trendline** object is a member of the **[Trendlines](Word.Trendlines.md)** collection. The **Trendlines** collection contains all the **Trendline** objects for a single series.


## Example

Use  **[Trendlines](Word.Series.Trendlines.md)** (Index), where Index is the trendline index number, to return a single **Trendline** object.

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


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]