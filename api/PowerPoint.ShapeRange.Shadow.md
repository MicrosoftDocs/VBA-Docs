---
title: ShapeRange.Shadow property (PowerPoint)
keywords: vbapp10.chm548033
f1_keywords:
- vbapp10.chm548033
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Shadow
ms.assetid: 01aa0a5a-341b-6764-e3ea-1f20379d0de3
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Shadow property (PowerPoint)

Returns a  **[ShadowFormat](PowerPoint.ShadowFormat.md)** object that contains shadow formatting properties for the specified shapes. Read-only.


## Syntax

_expression_.**Shadow**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


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


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]