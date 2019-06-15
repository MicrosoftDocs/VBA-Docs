---
title: TextStyle.Name property (Publisher)
keywords: vbapb10.chm5963782
f1_keywords:
- vbapb10.chm5963782
ms.prod: publisher
api_name:
- Publisher.TextStyle.Name
ms.assetid: 54e25e71-83d8-5074-fa0a-f956f075f482
ms.date: 06/15/2019
localization_priority: Normal
---


# TextStyle.Name property (Publisher)

Returns a **String** value indicating the name of the specified object. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[TextStyle](Publisher.TextStyle.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument.

For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]