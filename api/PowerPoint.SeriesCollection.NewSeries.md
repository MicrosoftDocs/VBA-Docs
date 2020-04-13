---
title: SeriesCollection.NewSeries method (PowerPoint)
keywords: vbapp10.chm66653
f1_keywords:
- vbapp10.chm66653
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesCollection.NewSeries
ms.assetid: 37a94558-02d9-7f0b-e881-0d9c5a9d4787
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection.NewSeries method (PowerPoint)

Creates a new series.


## Syntax

_expression_.**NewSeries**

_expression_ A variable that represents a '[SeriesCollection](PowerPoint.SeriesCollection.md)' object.


## Return value

A **[Series](PowerPoint.Series.md)** object that represents the new series.


## Remarks

This method is not available for PivotChart charts.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a new series to the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        Set ns = .Chart.SeriesCollection.NewSeries

    End If

End With
```


## See also


[SeriesCollection Object](PowerPoint.SeriesCollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]