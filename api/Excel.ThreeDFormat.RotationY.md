---
title: ThreeDFormat.RotationY property (Excel)
keywords: vbaxl10.chm119015
f1_keywords:
- vbaxl10.chm119015
ms.prod: excel
api_name:
- Excel.ThreeDFormat.RotationY
ms.assetid: 71d6e255-eb1c-62bc-61f2-8b4f8be3ad6f
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.RotationY property (Excel)

Returns or sets the rotation of the extruded shape around the y-axis in degrees. Can be a value from -90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write **Single**.


## Syntax

_expression_.**RotationY**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Remarks

To set the rotation of the extruded shape around the x-axis, use the **[RotationX](Excel.ThreeDFormat.RotationX.md)** property. 

To set the rotation of the extruded shape around the z-axis, use the **[Rotation](Excel.Shape.Rotation.md)** property of the **Shape** object. 

To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the **[SetExtrusionDirection](Excel.ThreeDFormat.SetExtrusionDirection.md)** method.


## Example

This example adds three identical extruded ovals to _myDocument_ and sets their rotation around the y-axis to -30, 0, and 30 degrees, respectively.

```vb
Set myDocument = Worksheets(1) 
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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]