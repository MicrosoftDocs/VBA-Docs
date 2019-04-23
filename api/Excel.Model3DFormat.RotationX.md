---
title: Model3DFormat.RotationX property (Excel)
ms.prod: excel
api_name:
- Excel.Model3DFormat.RotationX
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.RotationX property (Excel)

Returns the x-angle of a 3D model object's rotation. Read/write.

## Syntax

_expression_.**RotationX**

_expression_ A variable that represents a **[Model3DFormat](Excel.Model3DFormat.md)** object.


## Return value

Single

## Remarks

The rotation of a 3D model is reported as a trio of x, y, and z Euler angles. The properties **RotationX**, **[RotationY](Excel.Model3DFormat.RotationY.md)**, and **[RotationZ](Excel.Model3DFormat.RotationZ.md)** can be used to read or change the absolute orientation of a 3D model.  

The methods **[IncrementRotationX](Excel.Model3DFormat.IncrementRotationX.md)**, **[IncrementRotationY](Excel.Model3DFormat.IncrementRotationY.md)**, and **[IncrementRotationZ](Excel.Model3DFormat.IncrementRotationZ.md)** can be used to rotate a 3D model relative to its current orientation.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]