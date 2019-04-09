---
title: Model3DFormat.IncrementRotationX Method (Word)
keywords: vbawd10.chm151584880
f1_keywords:
- vbawd10.chm151584880
ms.prod: word
api_name:
- Word.Model3DFormat.IncrementRotationX
ms.date: 04/01/2019
localization_priority: Normal
---


# Model3DFormat.IncrementRotationX Method (Word)

Changes the rotation of the specified shape around the x-axis by the specified number of degrees. 


## Syntax

 _expression_.**IncrementRotationX** ( _Increment_ )

 _expression_ A variable that represents an [Model3DFormat](./Word.Model3DFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how much (in degrees) the rotation of the model around the x-axis is to be changed. Any value can be provided, although any value will be effectively normalized into the range 0..360 degrees.|

## Remarks

Use the  **[RotationX](Word.Model3DFormat.RotationX.md)** property to set the absolute rotation of the shape around the x-axis.

To change the rotation of a model around the y-axis, use the  **[IncrementRotationY](Word.Model3DFormat.IncrementRotationY.md)** method. To change the rotation around the z-axis, use the **[IncrementRotationZ](Word.Shape.IncrementRotationZ.md)** method.


## Example

This example tilts a 3D model on `myDocument` by 10 degrees. Shape one must be a 3D model for this code to have any effect.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Model3D.IncrementRotationX 10
```


## See also


[Model3DFormat Object](Word.Model3DFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]