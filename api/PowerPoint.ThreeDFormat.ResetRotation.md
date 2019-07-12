---
title: ThreeDFormat.ResetRotation method (PowerPoint)
keywords: vbapp10.chm557004
f1_keywords:
- vbapp10.chm557004
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.ResetRotation
ms.assetid: a766095a-f7a4-0fdf-8533-3ed00755950f
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.ResetRotation method (PowerPoint)

Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.


## Syntax

_expression_. `ResetRotation`

_expression_ A variable that represents a [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Remarks

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0 (zero), use the [RotationX](PowerPoint.ThreeDFormat.RotationX.md)and  **[RotationY](PowerPoint.ThreeDFormat.RotationY.md)** properties of the **ThreeDFormat** object. To set the extrusion rotation around the z-axis, use the **[Rotation](PowerPoint.Shape.Rotation.md)** property of the **[Shape](PowerPoint.Shape.md)** object that represents the extruded shape.


## Example

This example resets the rotation around the x-axis and the y-axis to 0 (zero) for the extrusion of shape one on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).ThreeD.ResetRotation
```


## See also


[ThreeDFormat Object](PowerPoint.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]