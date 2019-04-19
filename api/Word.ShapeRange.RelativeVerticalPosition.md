---
title: ShapeRange.RelativeVerticalPosition property (Word)
keywords: vbawd10.chm162857261
f1_keywords:
- vbawd10.chm162857261
ms.prod: word
api_name:
- Word.ShapeRange.RelativeVerticalPosition
ms.assetid: 4bcb0d85-53aa-e16d-98f3-4154de5355d8
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.RelativeVerticalPosition property (Word)

Specifies the relative vertical position of a range of shapes. Read/write **[WdRelativeHorizontalPosition](Word.WdRelativeHorizontalPosition.md)**.


## Syntax

_expression_. `RelativeVerticalPosition`

 _expression_ An expression that represents a **[ShapeRange](Word.shaperange.md)** object.


## Example

This example repositions the selected shape object.


```vb
With Selection.ShapeRange 
 .Left = InchesToPoints(0.6) 
 .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage 
 .Top = InchesToPoints(1) 
 .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph 
End With
```


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]