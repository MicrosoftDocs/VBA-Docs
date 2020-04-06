---
title: ChartGroup.SeriesCollection method (Word)
keywords: vbawd10.chm263454745
f1_keywords:
- vbawd10.chm263454745
ms.prod: word
api_name:
- Word.ChartGroup.SeriesCollection
ms.assetid: 4b4b7383-0967-cd2f-979c-eda9ef691459
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SeriesCollection method (Word)

Returns all the series in the chart group.


## Syntax

_expression_.**SeriesCollection** (_Index_)

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Return value

A  **[SeriesCollection](Word.SeriesCollection.md)** object that represents all the series in the chart group.


## Example

The following example turns on data labels for the first series of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1). _ 
 SeriesCollection(1).HasDataLabels = True 
 End If 
End With 

```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]