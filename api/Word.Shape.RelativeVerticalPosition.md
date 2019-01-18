---
title: Shape.RelativeVerticalPosition property (Word)
keywords: vbawd10.chm161481005
f1_keywords:
- vbawd10.chm161481005
ms.prod: word
api_name:
- Word.Shape.RelativeVerticalPosition
ms.assetid: 7e77dcab-5d1a-f955-1c80-2d130566167f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.RelativeVerticalPosition property (Word)

Specifies the relative vertical position of a shape. Read/write  **WdRelativeVerticalPosition**.


## Syntax

 _expression_. `RelativeVerticalPosition`

 _expression_ A variable that represents a '[Shape](Word.Shape.md)' object.


## Example

This example repositions the first shape object in the active document.


```vb
With ActiveDocument.Shapes(1) 
 .Left = InchesToPoints(0.6) 
 .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage 
 .Top = InchesToPoints(1) 
 .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph 
End With
```


## See also


[Shape Object](Word.Shape.md)

