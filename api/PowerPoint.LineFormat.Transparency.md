---
title: LineFormat.Transparency property (PowerPoint)
keywords: vbapp10.chm553013
f1_keywords:
- vbapp10.chm553013
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.Transparency
ms.assetid: 7d9e3a3c-479a-1a7a-45b2-4245b8444c21
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.Transparency property (PowerPoint)

Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0 (opaque) and 1.0 (clear). Read/write.


## Syntax

_expression_.**Transparency**

_expression_ A variable that represents a [LineFormat](PowerPoint.LineFormat.md) object.


## Return value

Single


## Remarks

The value of this property affects the appearance of solid-colored fills and lines only; it has no effect on the appearance of patterned lines or patterned, gradient, picture, or textured fills.


## Example

This example sets the shadow for shape three on _myDocument_ to semitransparent red. If the shape doesn't already have a shadow, this example adds one to it.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Shadow

    .Visible = True

    .ForeColor.RGB = RGB(255, 0, 0)

    .Transparency = 0.5

End With
```


## See also


[LineFormat Object](PowerPoint.LineFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]