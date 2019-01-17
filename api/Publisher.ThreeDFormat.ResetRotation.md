---
title: ThreeDFormat.ResetRotation Method (Publisher)
keywords: vbapb10.chm3801106
f1_keywords:
- vbapb10.chm3801106
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.ResetRotation
ms.assetid: 91e3943a-0087-fcb9-e33f-d41b60b869a7
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.ResetRotation Method (Publisher)

Resets the extrusion rotation around the x-axis (horizontal) and the y-axis (vertical) to 0 (zero) so that the front of the extrusion faces forward.


## Syntax

 _expression_. **ResetRotation**

 _expression_ A variable that represents a  **ThreeDFormat** object.


## Remarks

This method doesn't reset the rotation around the z-axis (extends outward from the plane of the publication).

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0, use the  **[RotationX](Publisher.ThreeDFormat.RotationX.md)** and  **[RotationY](Publisher.ThreeDFormat.RotationY.md)** properties of the  **ThreeDFormat** object.

To set the extrusion rotation around the z-axis, use the  **[Rotation](Publisher.Shape.Rotation.md)** property of the  **[Shape](Publisher.Shape.md)** object that represents the extruded shape.


## Example

This example resets the rotation around the x-axis and the y-axis to zero for the extrusion of the first shape in the active publication.


```vb
ActiveDocument.Pages(1).Shapes(1).ThreeD _ 
 .ResetRotation
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]