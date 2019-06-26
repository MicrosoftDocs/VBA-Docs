---
title: ThreeDFormat.ResetRotation method (Word)
keywords: vbawd10.chm164626444
f1_keywords:
- vbawd10.chm164626444
ms.prod: word
api_name:
- Word.ThreeDFormat.ResetRotation
ms.assetid: ab8b1bb6-2d39-2488-5db9-8405f8494014
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.ResetRotation method (Word)

Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward.


## Syntax

_expression_. `ResetRotation`

_expression_ Required. A variable that represents a '[ThreeDFormat](Word.ThreeDFormat.md)' object.


## Remarks

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0 (zero), use the  **RotationX** and **RotationY** properties of the **ThreeDFormat** object. To set the extrusion rotation around the z-axis, use the **Rotation** property of the **Shape** object that represents the extruded shape.


> [!NOTE] 
> This method does not reset the rotation around the z-axis.


## Example

This example resets the rotation around the x-axis and the y-axis to 0 (zero) for the extrusion of the first shape on the active document.


```vb
ActiveDocument.Shapes(1).ThreeD.ResetRotation
```


## See also


[ThreeDFormat Object](Word.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]