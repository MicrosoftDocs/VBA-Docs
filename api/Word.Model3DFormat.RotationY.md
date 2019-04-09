---
title: Model3DFormat.RotationY Property (Word)
keywords: vbawd10.chm151584870
f1_keywords:
- vbawd10.chm151584870
ms.prod: word
api_name:
- Word.Model3DFormat.RotationY
ms.date: 04/01/2019
localization_priority: Normal
---


# Model3DFormat.RotationY Property (Word)

Returns the y-angle of a 3D model object's rotation. Read/write.


## Return Value

Single


## Syntax

 _expression_.**RotationY**

 _expression_ A variable that represents a [Model3DFormat](./Word.Model3DFormat.md) object.


## Remarks

The rotation of a 3D model is reported as a trio of x, y, and z Euler angles.  Properties [RotationX](Word.Model3DFormat.RotationX.md), [RotationY](Word.Model3DFormat.RotationY.md), and [RotationZ](Word.Model3DFormat.RotationZ.md) can be used to read or change the absolute orientation of a 3D model.  Methods [IncrementRotationX](Word.Model3DFormat.IncrementRotationX.md), [IncrementRotationY](Word.Model3DFormat.IncrementRotationY.md), and [IncrementRotationZ](Word.Model3DFormat.IncrementRotationZ.md) can be used to rotate a 3D model relative to its current orientation.


## See also


[Model3DFormat Object](Word.Model3DFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]