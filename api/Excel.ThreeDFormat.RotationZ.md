---
title: ThreeDFormat.RotationZ property (Excel)
ms.prod: excel
api_name:
- Excel.ThreeDFormat.RotationZ
ms.assetid: 4e28396e-9d1e-4d2c-920c-e49e735cee27
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.RotationZ property (Excel)

Returns or sets the rotation of the extruded shape around the z-axis in degrees. Read/write  **Single**.


## Syntax

_expression_. `RotationZ`

_expression_ A variable that represents a [ThreeDFormat](./Excel.ThreeDFormat.md) object.


## Remarks

The  **RotationZ** property can be a value from - 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation.

To set the rotation of the extruded shape around the y-axis, use the  **RotationY** property of the ThreeDFormat object. To set the rotation of the extruded shape around the x-axis, use the **RotationX** property of the **Shape** object. To change the direction of the extrusion's sweep path without rotating the front face of the extrusion, use the **SetExtrusionDirection** method.


## See also


[ThreeDFormat Object](Excel.ThreeDFormat.md)

