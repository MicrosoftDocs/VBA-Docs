---
title: ShapeRange.Fill property (PowerPoint)
keywords: vbapp10.chm548022
f1_keywords:
- vbapp10.chm548022
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Fill
ms.assetid: 689cef96-6ad8-aa20-27c6-065af06b5753
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Fill property (PowerPoint)

Returns a **[FillFormat](PowerPoint.FillFormat.md)** object that contains fill formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**Fill**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

FillFormat


## Example

This example adds a rectangle to _myDocument_ and then sets the foreground color, background color, and gradient for the rectangle's fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
    .ForeColor.RGB = RGB(128, 0, 0)
    .BackColor.RGB = RGB(170, 170, 170)
    .TwoColorGradient msoGradientHorizontal, 1
End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]