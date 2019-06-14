---
title: ThreeDFormat.RotationX property (Publisher)
keywords: vbapb10.chm3801353
f1_keywords:
- vbapb10.chm3801353
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.RotationX
ms.assetid: 1ee394cb-746b-02f0-f2af-aa4a6fffd172
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.RotationX property (Publisher)

Returns or sets the rotation of the extruded shape around the x-axis in degrees. Can be a value from -90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write **Single**.


## Syntax

_expression_.**RotationX**

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Return value

Single


## Remarks

To set the rotation of the extruded shape around the y-axis, use the **[RotationY](Publisher.ThreeDFormat.RotationY.md)** property. 

To set the rotation of the extruded shape around the z-axis, use the **[Rotation](Publisher.Shape.Rotation.md)** property of the **Shape** object. 

To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the **[SetExtrusionDirection](Publisher.ThreeDFormat.SetExtrusionDirection.md)** method.


## Example

This example adds three identical extruded ovals to the active document and sets their rotation around the x-axis to -30, 0, and 30 degrees, respectively.

```vb
Sub SetRotationX() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShapeOval, Left:=30, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = -30 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=90, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = 0 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=150, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = 30 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]