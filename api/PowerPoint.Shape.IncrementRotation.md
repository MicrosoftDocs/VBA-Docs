---
title: Shape.IncrementRotation Method (PowerPoint)
keywords: vbapp10.chm547006
f1_keywords:
- vbapp10.chm547006
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.IncrementRotation
ms.assetid: f6e494fa-6bc1-0fc1-3bd3-ecc4fa5852e0
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.IncrementRotation Method (PowerPoint)

Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the  **Rotation** property to set the absolute rotation of the shape.


## Syntax

 _expression_. `IncrementRotation`( `_Increment_` )

 _expression_ A variable that represents a [Shape](./PowerPoint.Shape.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**| Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.|

## Remarks

To rotate a three-dimensional shape around the x-axis or the y-axis, use the  **[IncrementRotationX](PowerPoint.ThreeDFormat.IncrementRotationX.md)** method or the **[IncrementRotationY](PowerPoint.ThreeDFormat.IncrementRotationY.md)** method.


## Example

This example duplicates shape one on  `myDocument`, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Duplicate

    .Fill.PresetTextured msoTextureGranite

    .IncrementLeft 70

    .IncrementTop -50

    .IncrementRotation 30

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]