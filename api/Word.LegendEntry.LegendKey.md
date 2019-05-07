---
title: LegendEntry.LegendKey property (Word)
keywords: vbawd10.chm4784302
f1_keywords:
- vbawd10.chm4784302
ms.prod: word
api_name:
- Word.LegendEntry.LegendKey
ms.assetid: 11aa8dfa-fdb9-d7f1-3c03-17ce68dcdbec
ms.date: 06/08/2017
localization_priority: Normal
---


# LegendEntry.LegendKey property (Word)

Returns the legend key that is associated with the entry. Read-only  **[LegendKey](Word.LegendKey.md)**.


## Syntax

_expression_. `LegendKey`

_expression_ A variable that represents a '[LegendEntry](Word.LegendEntry.md)' object.


## Example

The following example sets the legend key for legend entry one on the first chart in the active document to be a triangle. You should run the example on a 2D line chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.LegendEntries(1).LegendKey _ 
 .MarkerStyle = xlMarkerStyleTriangle 
 End If 
End With
```


## See also


[LegendEntry Object](Word.LegendEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]