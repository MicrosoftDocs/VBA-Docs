---
title: InlineShape.Fill property (Word)
keywords: vbawd10.chm162005099
f1_keywords:
- vbawd10.chm162005099
ms.prod: word
api_name:
- Word.InlineShape.Fill
ms.assetid: d803d3cf-095f-a545-453d-4747a6e056c7
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.Fill property (Word)

Returns a  **[FillFormat](Word.FillFormat.md)** object that contains fill formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. `Fill`

 _expression_ A variable that represents an '[InlineShape](Word.InlineShape.md)' object.


## Example

This example adds a rectangle to myDocument and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = Documents(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 90, 90, 90, 50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]