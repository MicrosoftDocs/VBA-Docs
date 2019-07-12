---
title: Design.Name property (PowerPoint)
keywords: vbapp10.chm644008
f1_keywords:
- vbapp10.chm644008
ms.prod: powerpoint
api_name:
- PowerPoint.Design.Name
ms.assetid: a851e05b-9697-0f84-be62-a968e423f74a
ms.date: 06/08/2017
localization_priority: Normal
---


# Design.Name property (PowerPoint)

Returns or sets the name of the specified object. Read/write.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Design](PowerPoint.Design.md) object.


## Return value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.


## See also


[Design Object](PowerPoint.Design.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]