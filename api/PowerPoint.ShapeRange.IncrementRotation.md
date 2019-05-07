---
title: ShapeRange.IncrementRotation method (PowerPoint)
keywords: vbapp10.chm548006
f1_keywords:
- vbapp10.chm548006
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.IncrementRotation
ms.assetid: 427367bb-5264-86de-cf39-be252c4b7098
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.IncrementRotation method (PowerPoint)

Changes the rotation of the specified shape range around the z-axis by the specified number of degrees. Use the  **Rotation** property to set the absolute rotation of the shape range.


## Syntax

_expression_. `IncrementRotation`( `_Increment_` )

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how far the shape range is to be rotated horizontally, in degrees. A positive value rotates the shape range clockwise; a negative value rotates it counterclockwise.|

## Remarks

To rotate a three-dimensional shape range around the x-axis or the y-axis, use the  **[IncrementRotationX](PowerPoint.ThreeDFormat.IncrementRotationX.md)** method or the **[IncrementRotationY](PowerPoint.ThreeDFormat.IncrementRotationY.md)** method.


## Example

This example duplicates shape one on _myDocument_, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


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


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]