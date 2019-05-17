---
title: ThreeDFormat.IncrementRotationY method (Excel)
keywords: vbaxl10.chm119021
f1_keywords:
- vbaxl10.chm119021
ms.prod: excel
api_name:
- Excel.ThreeDFormat.IncrementRotationY
ms.assetid: 56dde624-a56d-41f1-3192-f4c5c28e0a66
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.IncrementRotationY method (Excel)

Changes the rotation of the specified shape around the y-axis by the specified number of degrees. Use the **[RotationY](Excel.ThreeDFormat.RotationY.md)** property to set the absolute rotation of the shape around the y-axis.


## Syntax

_expression_.**IncrementRotationY** (_Increment_)

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.|

## Remarks

You cannot adjust the specified shape's rotation around the y-axis past the upper or lower limit for the **RotationY** property (90 degrees to -90 degrees). For example, if the **RotationY** property is initially set to 80 and you specify 40 for the _Increment_ argument, the resulting rotation will be 90 (the upper limit for the **RotationY** property) instead of 120.

To change the rotation of a shape around the x-axis, use the **[IncrementRotationX](Excel.ThreeDFormat.IncrementRotationX.md)** method. To change the rotation around the z-axis, use the **[IncrementRotationZ](Excel.ThreeDFormat.IncrementRotationZ.md)** method.


## Example

This example tilts shape one on _myDocument_ 10 degrees to the right. Shape one must be an extruded shape for you to see the effect of this code.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).ThreeD.IncrementRotationY -10
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]