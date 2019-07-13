---
title: ThreeDFormat.IncrementRotationY method (PowerPoint)
keywords: vbapp10.chm557003
f1_keywords:
- vbapp10.chm557003
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.IncrementRotationY
ms.assetid: a9216bbc-8e82-4b6d-d1ac-f67a4a44a092
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.IncrementRotationY method (PowerPoint)

Changes the rotation of the specified shape around the y-axis by the specified number of degrees. 


## Syntax

_expression_. `IncrementRotationY`( `_Increment_` )

_expression_ A variable that represents an [ThreeDFormat](PowerPoint.ThreeDFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from -90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.|

## Remarks

Use the  **[RotationY](PowerPoint.ThreeDFormat.RotationY.md)** property to set the absolute rotation of the shape around the y-axis.

To change the rotation of a shape around the x-axis, use the  **[IncrementRotationX](PowerPoint.ThreeDFormat.IncrementRotationX.md)** method. To change the rotation around the z-axis, use the **[IncrementRotation](PowerPoint.Shape.IncrementRotation.md)** method.

You cannot adjust the rotation around the y-axis of the specified shape past the upper or lower limit for the  **RotationY** property (90 degrees to -90 degrees). For example, if the **RotationY** property is initially set to 80 and you specify 40 for the Increment argument, the resulting rotation will be 90 (the upper limit for the **RotationY** property) instead of 120.


## Example

This example tilts shape one on _myDocument_ 10 degrees to the right. Shape one must be an extruded shape for you to see the effect of this code.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).ThreeD.IncrementRotationY -10
```


## See also


[ThreeDFormat Object](PowerPoint.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]