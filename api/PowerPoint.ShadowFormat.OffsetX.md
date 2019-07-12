---
title: ShadowFormat.OffsetX property (PowerPoint)
keywords: vbapp10.chm554006
f1_keywords:
- vbapp10.chm554006
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.OffsetX
ms.assetid: 54c43556-99cc-dbd4-5192-abd57798d73f
ms.date: 06/08/2017
localization_priority: Normal
---


# ShadowFormat.OffsetX property (PowerPoint)

Returns or sets the horizontal offset of the shadow from the specified shape, in points. Read/write.


## Syntax

_expression_.**OffsetX**

_expression_ A variable that represents an [ShadowFormat](PowerPoint.ShadowFormat.md) object.


## Return value

Single


## Remarks

A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.

If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the  **[IncrementOffsetX](PowerPoint.ShadowFormat.IncrementOffsetX.md)** method or the **[IncrementOffsetY](PowerPoint.ShadowFormat.IncrementOffsetY.md)** method.


## Example

This example sets the horizontal and vertical offsets of the shadow for shape three on _myDocument_. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Shadow

    .Visible = True

    .OffsetX = 5

    .OffsetY = -3

End With
```


## See also


[ShadowFormat Object](PowerPoint.ShadowFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]