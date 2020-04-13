---
title: ThreeDFormat.RotationY property (Word)
keywords: vbawd10.chm164626542
f1_keywords:
- vbawd10.chm164626542
ms.prod: word
api_name:
- Word.ThreeDFormat.RotationY
ms.assetid: 64ebb9d9-4338-3672-9149-e1e82ba8abdc
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.RotationY property (Word)

Returns or sets the rotation of the extruded shape around the y-axis, in degrees. Read/write  **Single**.


## Syntax

_expression_. `RotationY`

 _expression_ An expression that returns a '[ThreeDFormat](Word.ThreeDFormat.md)' object.


## Remarks

The **RotationY** property can be a value from - 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right.

To set the rotation of the extruded shape around the x-axis, use the **[RotationX](Word.ThreeDFormat.RotationX.md)** property of the ThreeDFormat object. To set the rotation of the extruded shape around the z-axis, use the **[Rotation](Word.Shape.Rotation.md)** property of the **[Shape](Word.Shape.md)** object. To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the **[SetExtrusionDirection](Word.ThreeDFormat.SetExtrusionDirection.md)** method.


## Example

This example adds three identical extruded ovals to myDocument and sets their rotation around the y-axis to - 30, 0, and 30 degrees, respectively.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes 
 With .AddShape(msoShapeOval, 30, 30, 50, 25).ThreeD 
 .Visible = True 
 .RotationY = -30 
 End With 
 With .AddShape(msoShapeOval, 30, 70, 50, 25).ThreeD 
 .Visible = True 
 .RotationY = 0 
 End With 
 With .AddShape(msoShapeOval, 30, 110, 50, 25).ThreeD 
 .Visible = True 
 .RotationY = 30 
 End With 
End With
```


## See also


[ThreeDFormat Object](Word.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]