---
title: Axis.TickLabels property (Word)
keywords: vbawd10.chm113049650
f1_keywords:
- vbawd10.chm113049650
ms.prod: word
api_name:
- Word.Axis.TickLabels
ms.assetid: 5c363e25-71e3-4f89-bcd3-612855000f53
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.TickLabels property (Word)

Returns the tick-mark labels for the specified axis. Read-only  **[TickLabels](Word.TickLabels.md)**.


## Syntax

_expression_. `TickLabels`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Example

The following example sets the color of the tick-mark label font for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).TickLabels.Font.ColorIndex = 3 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]