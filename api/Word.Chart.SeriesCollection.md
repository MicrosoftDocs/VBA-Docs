---
title: Chart.SeriesCollection method (Word)
ms.prod: word
api_name:
- Word.Chart.SeriesCollection
ms.assetid: b9688aef-839a-b45b-1596-d8f02225aa05
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SeriesCollection method (Word)

Returns all the series in the chart.


## Syntax

_expression_.**SeriesCollection** (_Index_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Return value

A  **[SeriesCollection](Word.SeriesCollection.md)** object that represents all the series in the chart.


## Example

The following example turns on data labels for series one of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).HasDataLabels = True 
 End If 
End With 

```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]