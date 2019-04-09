---
title: Model3DFormat.RotationZ property (PowerPoint)
keywords: vbapp10.chm743012
f1_keywords:
- vbapp10.chm743012
ms.prod: powerpoint
api_name:
- PowerPoint.Model3DFormat.RotationZ
ms.date: 04/01/2019
localization_priority: Normal
---


# Model3DFormat.RotationZ property (PowerPoint)

Returns the z-angle of a 3D model object's rotation. Read/write.


## Return value

Single


## Syntax

_expression_.**RotationZ**

 _expression_ A variable that represents a [Model3DFormat](./PowerPoint.Model3DFormat.md) object.


## Remarks

The rotation of a 3D model is reported as a trio of x, y, and z Euler angles.  Properties [RotationX](PowerPoint.Model3DFormat.RotationX.md), [RotationY](PowerPoint.Model3DFormat.RotationY.md), and [RotationZ](PowerPoint.Model3DFormat.RotationZ.md) can be used to read or change the absolute orientation of a 3D model.  Methods [IncrementRotationX](PowerPoint.Model3DFormat.IncrementRotationX.md), [IncrementRotationY](PowerPoint.Model3DFormat.IncrementRotationY.md), and [IncrementRotationZ](PowerPoint.Model3DFormat.IncrementRotationZ.md) can be used to rotate a 3D model relative to its current orientation.


## See also


[Model3DFormat Object](PowerPoint.Model3DFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]