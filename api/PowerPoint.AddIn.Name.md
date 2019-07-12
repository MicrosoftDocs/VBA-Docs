---
title: AddIn.Name property (PowerPoint)
keywords: vbapp10.chm521004
f1_keywords:
- vbapp10.chm521004
ms.prod: powerpoint
api_name:
- PowerPoint.AddIn.Name
ms.assetid: d5a859ab-9304-1148-315d-2d2983251197
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIn.Name property (PowerPoint)

The name (title) of the add-in for file types that are registered. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents an [AddIn](PowerPoint.AddIn.md) object.


## Return value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.


## See also


[AddIn Object](PowerPoint.AddIn.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]