---
title: Shape.PickUp method (PowerPoint)
keywords: vbapp10.chm547008
f1_keywords:
- vbapp10.chm547008
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.PickUp
ms.assetid: 35730a7e-3878-dfae-2aba-3395d41e5f3e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.PickUp method (PowerPoint)

Copies the formatting of the specified shape. Use the  **Apply** method to apply the copied formatting to another shape.


## Syntax

_expression_.**PickUp**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


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


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]