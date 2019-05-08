---
title: Frame.HorizontalDistanceFromText property (Word)
keywords: vbawd10.chm153747459
f1_keywords:
- vbawd10.chm153747459
ms.prod: word
api_name:
- Word.Frame.HorizontalDistanceFromText
ms.assetid: 40672084-cced-8807-8843-0750ef5a48b9
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame.HorizontalDistanceFromText property (Word)

Returns or sets the horizontal distance between a frame and the surrounding text, in points. Read/write  **Single**.


## Syntax

_expression_. `HorizontalDistanceFromText`

_expression_ A variable that represents a '[Frame](Word.Frame.md)' object.


## Example

This example adds a frame around the selection and sets the horizontal distance between the frame and the text to 12 points.


```vb
Dim frmNew As Frame 
 
Set frmNew = ActiveDocument.Frames.Add(Range:=Selection.Range) 
frmNew.HorizontalDistanceFromText = 12
```

This example adds a frame around the first paragraph and sets several properties of the frame.




```vb
Dim frmNew As Frame 
 
Set frmNew = ActiveDocument.Frames.Add _ 
 (Range:=ActiveDocument.Paragraphs(1).Range) 
 
With frmNew 
 .HorizontalDistanceFromText = InchesToPoints(0.25) 
 .VerticalDistanceFromText = InchesToPoints(0.25) 
 .HeightRule = wdFrameAuto 
 .WidthRule = wdFrameAuto 
 .Borders.Enable = False 
End With
```


## See also


[Frame Object](Word.Frame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]