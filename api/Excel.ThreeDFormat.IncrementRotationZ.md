---
title: ThreeDFormat.IncrementRotationZ method (Excel)
ms.prod: excel
api_name:
- Excel.ThreeDFormat.IncrementRotationZ
ms.assetid: 3301f928-81d4-3dba-121a-18c0a8aeef5f
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.IncrementRotationZ method (Excel)

Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the **[RotationZ](Excel.ThreeDFormat.RotationZ.md)** property to set the absolute rotation of the shape around the z-axis.


## Syntax

_expression_.**IncrementRotationZ** (_Increment_)

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much (in degrees) the rotation of the shape around the z-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.|

## Remarks

You cannot adjust the specified shape's rotation around the z-axis past the upper or lower limit for the **RotationZ** property (90 degrees to -90 degrees). For example, if the **RotationZ** property is initially set to 80 and you specify 40 for the _Increment_ argument, the resulting rotation will be 90 (the upper limit for the **RotationZ** property) instead of 120.

To change the rotation of a shape around the x-axis, use the **[IncrementRotationX](Excel.ThreeDFormat.IncrementRotationX.md)** method. To change the rotation of a shape around the y-axis, use the **[IncrementRotationY](Excel.ThreeDFormat.IncrementRotationY.md)** method. 



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]