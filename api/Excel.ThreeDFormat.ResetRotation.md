---
title: ThreeDFormat.ResetRotation method (Excel)
keywords: vbaxl10.chm119002
f1_keywords:
- vbaxl10.chm119002
ms.prod: excel
api_name:
- Excel.ThreeDFormat.ResetRotation
ms.assetid: 55173d20-2d13-d3a8-39db-6b1a161c6ea6
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.ResetRotation method (Excel)

Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.


## Syntax

_expression_.**ResetRotation**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Remarks

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0 (zero), use the **[RotationX](Excel.ThreeDFormat.RotationX.md)** and **[RotationY](Excel.ThreeDFormat.RotationY.md)** properties of the **ThreeDFormat** object. 

To set the extrusion rotation around the z-axis, use the **[Rotation](Excel.Shape.Rotation.md)** property of the **Shape** object that represents the extruded shape.


## Example

This example resets the rotation around the x-axis and the y-axis to 0 (zero) for the extrusion of shape one on _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).ThreeD.ResetRotation
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]