---
title: Frame.RelativeVerticalPosition property (Word)
keywords: vbawd10.chm153747464
f1_keywords:
- vbawd10.chm153747464
ms.prod: word
api_name:
- Word.Frame.RelativeVerticalPosition
ms.assetid: 70da43d6-602b-3afc-3353-a4ac53a48534
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame.RelativeVerticalPosition property (Word)

Specifies the relative vertical position of a frame. Read/write  **[WdRelativeVerticalPosition](Word.WdRelativeVerticalPosition.md)**.


## Syntax

_expression_. `RelativeVerticalPosition`

_expression_ A variable that represents a '[Frame](Word.Frame.md)' object.


## Example

This example adds a frame around the selection and aligns the frame vertically with the top of the page.


```vb
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range) 
With myFrame 
 .RelativeVerticalPosition = wdRelativeVerticalPositionPage 
 .VerticalPosition = wdFrameTop 
End With
```


## See also


[Frame Object](Word.Frame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]