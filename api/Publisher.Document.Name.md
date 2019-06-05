---
title: Document.Name property (Publisher)
keywords: vbapb10.chm196630
f1_keywords:
- vbapb10.chm196630
ms.prod: publisher
api_name:
- Publisher.Document.Name
ms.assetid: fcf86fcc-a3aa-b4c6-1ecc-202972ac558b
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.Name property (Publisher)

Returns a **String** value indicating the name of the specified object. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]