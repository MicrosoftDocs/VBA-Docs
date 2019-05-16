---
title: Shape.ZOrderPosition property (Excel)
keywords: vbaxl10.chm636116
f1_keywords:
- vbaxl10.chm636116
ms.prod: excel
api_name:
- Excel.Shape.ZOrderPosition
ms.assetid: aaf86516-bf5d-bdb5-1d88-eb1784f9b26f
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.ZOrderPosition property (Excel)

Returns the position of the specified shape in the z-order. Read-only **Long**.


## Syntax

_expression_.**ZOrderPosition**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Remarks

To set the shape's position in the z-order, use the **[ZOrder](Excel.Shape.ZOrder.md)** method.

A shape's position in the z-order corresponds to the shape's index number in the **[Shapes](Excel.Shapes.md)** collection. For example, if there are four shapes on _myDocument_, the expression  `myDocument.Shapes(1)` returns the shape at the back of the z-order, and the expression `myDocument.Shapes(4)` returns the shape at the front of the z-order.

Whenever you add a new shape to a collection, it's added to the front of the z-order by default.


## Example

This example adds an oval to _myDocument_ and then places the oval second from the back in the z-order if there is at least one other shape on the document.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300) 
 While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Wend 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]