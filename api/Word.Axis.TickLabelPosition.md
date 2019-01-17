---
title: Axis.TickLabelPosition property (Word)
keywords: vbawd10.chm113049648
f1_keywords:
- vbawd10.chm113049648
ms.prod: word
api_name:
- Word.Axis.TickLabelPosition
ms.assetid: c0284fd9-ec02-fdc9-4c8b-49efdb85be87
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.TickLabelPosition property (Word)

Describes the position of tick-mark labels on the specified axis. Read/write  **[xlTickLabelPosition](Word.xlticklabelposition.md)**.


## Syntax

 _expression_. `TickLabelPosition`

 _expression_ A variable that represents an '[Axis](Word.Axis.md)' object.


## Example

The following example sets tick-mark labels to the high position (above the chart) on the category axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory) _ 
 .TickLabelPosition = xlTickLabelPositionHigh 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

