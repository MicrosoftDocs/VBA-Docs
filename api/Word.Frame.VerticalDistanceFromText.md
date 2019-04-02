---
title: Frame.VerticalDistanceFromText property (Word)
keywords: vbawd10.chm153747465
f1_keywords:
- vbawd10.chm153747465
ms.prod: word
api_name:
- Word.Frame.VerticalDistanceFromText
ms.assetid: 16496bd5-bb12-11ad-59e6-baf234803471
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame.VerticalDistanceFromText property (Word)

Returns or sets the vertical distance (in points) between a frame and the surrounding text. Read/write  **Single**.


## Syntax

_expression_. `VerticalDistanceFromText`

 _expression_ An expression that returns a '[Frame](Word.Frame.md)' object.


## Example

This example sets the vertical distance between the selected frame and the surrounding text to 12 points.


```vb
If Selection.Frames.Count = 1 Then 
 Selection.Frames(1).VerticalDistanceFromText = 12 
End If
```

This example adds a frame around the selection and sets several properties of the frame.




```vb
Set aFrame = ActiveDocument.Frames.Add(Range:=Selection.Range) 
With aFrame 
 .HorizontalDistanceFromText = InchesToPoints(0.13) 
 .VerticalDistanceFromText = InchesToPoints(0.13) 
 .HeightRule = wdFrameAuto 
 .WidthRule = wdFrameAuto 
End With
```


## See also


[Frame Object](Word.Frame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]