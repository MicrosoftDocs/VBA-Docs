---
title: Model3DFormat.IncrementRotationY method (Word)
keywords: vbawd10.chm151584881
f1_keywords:
- vbawd10.chm151584881
ms.prod: word
api_name:
- Word.Model3DFormat.IncrementRotationY
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.IncrementRotationY method (Word)

Changes the rotation of the specified shape around the y-axis by the specified number of degrees. 


## Syntax

_expression_.**IncrementRotationY** (_Increment_)

_expression_ A variable that represents a **[Model3DFormat](Word.Model3DFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how much (in degrees) the rotation of the model around the y-axis is to be changed. Any value can be provided, although any value will be effectively normalized into the range 0..360 degrees.|

## Remarks

Use the **[RotationY](Word.Model3DFormat.RotationY.md)** property to set the absolute rotation of the shape around the y-axis.

To change the rotation of a model around the x-axis, use the **[IncrementRotationX](Word.Model3DFormat.IncrementRotationX.md)** method. To change the rotation around the z-axis, use the **[IncrementRotationZ](Word.Model3DFormat.IncrementRotationZ.md)** method.


## Example

This example tilts a 3D model on _myDocument_ by 10 degrees. Shape one must be a 3D model for this code to have any effect.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Model3D.IncrementRotationY 10
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]