---
title: ShapeRange.IncrementRotation method (Excel)
keywords: vbaxl10.chm640084
f1_keywords:
- vbaxl10.chm640084
ms.prod: excel
api_name:
- Excel.ShapeRange.IncrementRotation
ms.assetid: 42da1be8-a858-d910-bda1-ed174dc7dd24
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.IncrementRotation method (Excel)

Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the **[Rotation](Excel.ShapeRange.Rotation.md)** property to set the absolute rotation of the shape.


## Syntax

_expression_.**IncrementRotation** (_Increment_)

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.|

## Remarks

To rotate a three-dimensional shape around the x-axis or the y-axis, use the **[IncrementRotationX](Excel.ThreeDFormat.IncrementRotationX.md)** method or the **[IncrementRotationY](Excel.ThreeDFormat.IncrementRotationY.md)** method.


## Example

This example duplicates shape one on _myDocument_, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]