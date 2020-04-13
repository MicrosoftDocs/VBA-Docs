---
title: Frame.HorizontalPosition property (Word)
keywords: vbawd10.chm153747461
f1_keywords:
- vbawd10.chm153747461
ms.prod: word
api_name:
- Word.Frame.HorizontalPosition
ms.assetid: e71b0df4-53c0-d917-b1b7-32b0ee5205aa
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame.HorizontalPosition property (Word)

Returns or sets the horizontal distance between the edge of the frame and the item specified by the **[RelativeHorizontalPosition](Word.Frame.RelativeHorizontalPosition.md)** property. Read/write **Single**.


## Syntax

_expression_. `HorizontalPosition`

_expression_ A variable that represents a '[Frame](Word.Frame.md)' object.


## Remarks

This property can be a number that indicates a measurement in points, or can be one of the **[WdFramePosition](Word.WdFramePosition.md)** constants.


## Example

This example aligns the first frame in the active document horizontally with the right margin.


```vb
If ActiveDocument.Frames.Count >= 1 Then 
 With ActiveDocument.Frames(1) 
 .RelativeHorizontalPosition = _ 
 wdRelativeHorizontalPositionMargin 
 .HorizontalPosition = wdFrameRight 
 End With 
End If
```


## See also


[Frame Object](Word.Frame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]