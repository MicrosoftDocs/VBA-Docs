---
title: Model3DFormat.RotationZ property (Word)
keywords: vbawd10.chm151584871
f1_keywords:
- vbawd10.chm151584871
ms.prod: word
api_name:
- Word.Model3DFormat.RotationZ
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.RotationZ property (Word)

Returns the z-angle of a 3D model object's rotation. Read/write.

## Syntax

_expression_.**RotationZ**

_expression_ A variable that represents a **[Model3DFormat](Word.Model3DFormat.md)** object.


## Return value

Single

## Remarks

The rotation of a 3D model is reported as a trio of x, y, and z Euler angles.  The properties **[RotationX](Word.Model3DFormat.RotationX.md)**, **[RotationY](Word.Model3DFormat.RotationY.md)**, and **RotationZ** can be used to read or change the absolute orientation of a 3D model.  

The methods **[IncrementRotationX](Word.Model3DFormat.IncrementRotationX.md)**, **[IncrementRotationY](Word.Model3DFormat.IncrementRotationY.md)**, and **[IncrementRotationZ](Word.Model3DFormat.IncrementRotationZ.md)** can be used to rotate a 3D model relative to its current orientation.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]