---
title: Shape.Shadow property (PowerPoint)
keywords: vbapp10.chm547033
f1_keywords:
- vbapp10.chm547033
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Shadow
ms.assetid: 832b8e62-4fc5-1f4b-74c7-cc0e63a12699
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Shadow property (PowerPoint)

Returns a **[ShadowFormat](PowerPoint.ShadowFormat.md)** object that contains shadow formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**Shadow**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Example

This example adds a shadowed rectangle to slide one in the active presentation. The blue, embossed shadow is offset 3 points to the right of and 2 points down from the rectangle.


```vb
Set myShape = Application.ActivePresentation.Slides(1).Shapes

With myShape.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow

    .Type = msoShadow17

    .ForeColor.RGB = RGB(0, 0, 128)

    .OffsetX = 3

    .OffsetY = 2

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]