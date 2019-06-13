---
title: Shape.IncrementRotation method (Publisher)
keywords: vbapb10.chm2228257
f1_keywords:
- vbapb10.chm2228257
ms.prod: publisher
api_name:
- Publisher.Shape.IncrementRotation
ms.assetid: 3293c707-f3e8-1afb-cf9c-231ceae66ab6
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.IncrementRotation method (Publisher)

Changes the rotation of the specified shape around the z-axis (extends outward from the plane of the publication) by the specified number of degrees.


## Syntax

_expression_.**IncrementRotation** (_Increment_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Increment_|Required| **Single**|Specifies how far the shape is to be rotated around the z-axis, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise. Valid values are between -360 and 360.|

## Return value

Nothing


## Remarks

Use the **[Rotation](Publisher.ShapeRange.Rotation.md)** property of the **ShapeRange** object to set the absolute rotation of the shape.

To rotate a three-dimensional shape around the x-axis (horizontal) or the y-axis (vertical), use the **[IncrementRotationX](Publisher.ThreeDFormat.IncrementRotationX.md)** method or the **[IncrementRotationY](Publisher.ThreeDFormat.IncrementRotationY.md)** method of the **ThreeDFormat** object, respectively.


## Example

This example duplicates the first shape on the active publication, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.

```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .Fill.PresetTextured PresetTexture:=msoTextureGranite 
 .IncrementLeft Increment:=70 
 .IncrementTop Increment:=-50 
 .IncrementRotation Increment:=30 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]