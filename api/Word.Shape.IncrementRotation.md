---
title: Shape.IncrementRotation method (Word)
keywords: vbawd10.chm161480719
f1_keywords:
- vbawd10.chm161480719
ms.prod: word
api_name:
- Word.Shape.IncrementRotation
ms.assetid: 67f44fb6-0cce-9a5c-5ac7-b8116dffc167
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.IncrementRotation method (Word)

Changes the rotation of the specified shape around the z-axis by the specified number of degrees.


## Syntax

_expression_. `IncrementRotation`( `_Increment_` )

_expression_ Required. A variable that represents a **[Shape](Word.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.|

## Remarks

Use the  **[Rotation](Word.Shape.Rotation.md)** property to set the absolute rotation of the shape. To rotate a three-dimensional shape around the x-axis or the y-axis, use the **[IncrementRotationX](Word.ThreeDFormat.IncrementRotationX.md)** or **[IncrementRotationY](Word.ThreeDFormat.IncrementRotationY.md)** method of the **[ThreeDFormat](Word.ThreeDFormat.md)** object.


## Example

This example duplicates shape one on _myDocument_ , sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]