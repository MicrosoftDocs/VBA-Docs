---
title: ChartFont.Color property (Word)
keywords: vbawd10.chm255918084
f1_keywords:
- vbawd10.chm255918084
ms.prod: word
api_name:
- Word.ChartFont.Color
ms.assetid: 8d5aebd9-975a-63a6-1c2f-930f588b4004
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Color property (Word)

Returns or sets the primary color of the object. Read/write  **Variant**.


## Syntax

_expression_.**Color**

_expression_ A variable that represents a '[ChartFont](Word.ChartFont.md)' object.


## Example

The following example sets the color of the tick-mark labels on the value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 Chart.Axes(xlValue).TickLabels.Font.Color = _ 
 RGB(0, 255, 0) 
 End If 
End With
```


## See also


[ChartFont Object](Word.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]