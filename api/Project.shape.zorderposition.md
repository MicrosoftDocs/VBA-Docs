---
title: Shape.ZOrderPosition property (Project)
ms.prod: project-server
ms.assetid: ebbd573a-4cf0-a3af-7dff-de67d321d9d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ZOrderPosition property (Project)
Gets the position of the shape in the z-order. Read-only  **Long**.

## Syntax

_expression_.**ZOrderPosition**

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Remarks

To set the shape position in the z-order, use the [ZOrder](Project.shape.zorder.md) method.

The position of a shape in the z-order corresponds to the index number of the shape in the  **Shapes** collection. For example, if there are four shapes in the `myReport` report object, the expression `myReport.Shapes(1)` returns the shape at the back of the z-order, and the expression `myReport.Shapes(4)` returns the shape at the front of the z-order.

When you add a shape to a  **Shapes** collection, the shape is added to the front of the z-order by default.


## Property value

 **INT**


## See also


[Shape Object](Project.shape.md)
[Shapes Object](Project.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]