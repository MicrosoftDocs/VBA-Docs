---
title: ShapeRange.PickUp method (PowerPoint)
keywords: vbapp10.chm548008
f1_keywords:
- vbapp10.chm548008
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.PickUp
ms.assetid: f583c44c-0ab1-19eb-40f7-7e3412c93686
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.PickUp method (PowerPoint)

Copies the formatting of the specified shape. Use the  **Apply** method to apply the copied formatting to another shape.


## Syntax

_expression_.**PickUp**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Example

This example copies the formatting of shape one on _myDocument_ and then applies the copied formatting to shape two.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument

    .Shapes(1).PickUp

    .Shapes(2).Apply

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]