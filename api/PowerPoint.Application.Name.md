---
title: Application.Name property (PowerPoint)
keywords: vbapp10.chm502009
f1_keywords:
- vbapp10.chm502009
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Name
ms.assetid: c7a59327-774a-8c55-17b4-053ae76bd623
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Name property (PowerPoint)

Returns the string "Microsoft PowerPoint." Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]