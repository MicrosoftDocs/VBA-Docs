---
title: Shape.Rotation property (PowerPoint)
keywords: vbapp10.chm547031
f1_keywords:
- vbapp10.chm547031
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Rotation
ms.assetid: 3ed090a8-d945-85ee-155b-08628feff348
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Rotation property (PowerPoint)

Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write.


## Syntax

_expression_.**Rotation**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

Single


## Remarks

A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. 

To set the rotation of a three-dimensional shape around the x-axis or the y-axis, use the  **[RotationX](PowerPoint.ThreeDFormat.RotationX.md)** property or the **[RotationY](PowerPoint.ThreeDFormat.RotationY.md)** property of the **[ThreeDFormat](PowerPoint.ThreeDFormat.md)** object.


## Example

This example matches the rotation of all shapes on _myDocument_ to the rotation of shape one.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    sh1Rotation = .Item(1).Rotation

    For o = 1 To .Count

        .Item(o).Rotation = sh1Rotation

    Next

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]