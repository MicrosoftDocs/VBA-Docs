---
title: Tag.Name property (Publisher)
keywords: vbapb10.chm4718595
f1_keywords:
- vbapb10.chm4718595
ms.prod: publisher
api_name:
- Publisher.Tag.Name
ms.assetid: a35e8c51-e4c8-2554-eb44-8f202795fbc7
ms.date: 06/15/2019
localization_priority: Normal
---


# Tag.Name property (Publisher)

Returns a **String** value indicating the name of the specified object. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Tag](Publisher.Tag.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. 

For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]