---
title: SeriesCollection.NewSeries method (Word)
keywords: vbawd10.chm150406237
f1_keywords:
- vbawd10.chm150406237
ms.prod: word
api_name:
- Word.SeriesCollection.NewSeries
ms.assetid: fbfe3d37-c099-508e-367d-27314dc5c8ae
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection.NewSeries method (Word)

Creates a new series.


## Syntax

_expression_.**NewSeries**

_expression_ A variable that represents a **[SeriesCollection](Word.SeriesCollection.md)** object.


## Return value

A  **[Series](Word.Series.md)** object that represents the new series.


## Remarks

This method is not available for PivotChart charts.


## Example

The following example adds a new series to the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 Set ns = .Chart.SeriesCollection.NewSeries 
 End If 
End With
```


## See also


[SeriesCollection Object](Word.SeriesCollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]