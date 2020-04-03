---
title: Legend.LegendEntries method (Word)
keywords: vbawd10.chm147194029
f1_keywords:
- vbawd10.chm147194029
ms.prod: word
api_name:
- Word.Legend.LegendEntries
ms.assetid: 4dc6b7bf-3a65-3080-17e0-eb58ffb978b0
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend.LegendEntries method (Word)

Returns a collection of legend entries for the legend.


## Syntax

_expression_. `LegendEntries`

_expression_ A variable that represents a '[Legend](Word.Legend.md)' object.


## Return value

A **[LegendEntries](Word.LegendEntries.md)** object that represents the legend entries for the legend.


## Example

The following example sets the font for legend entry one on the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.LegendEntries(1).Font.Name = "Arial" 
 End If 
End With
```


## See also


[Legend Object](Word.Legend.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]