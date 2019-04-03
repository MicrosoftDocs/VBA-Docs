---
title: Series object (Word)
keywords: vbawd10.chm1888
f1_keywords:
- vbawd10.chm1888
ms.prod: word
api_name:
- Word.Series
ms.assetid: 212c323f-8acb-2ba7-1359-ab0f43268e77
ms.date: 06/08/2017
localization_priority: Normal
---


# Series object (Word)

Represents a series in a chart.


## Remarks

 The **Series** object is a member of the **[SeriesCollection](Word.SeriesCollection.md)** collection.


## Example

Use  **[SeriesCollection](Word.Chart.SeriesCollection.md)** (_index_), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series of the first chart in the active document.

The series index number indicates the order in which the series were added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.




```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0) 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]