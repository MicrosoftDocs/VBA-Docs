---
title: Shape.Name property (Publisher)
keywords: vbapb10.chm2228292
f1_keywords:
- vbapb10.chm2228292
ms.prod: publisher
api_name:
- Publisher.Shape.Name
ms.assetid: 307c131b-f6ad-38e7-d214-420063d3e5ec
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Name property (Publisher)

Returns or sets a **String** value indicating the name of the specified object. Read/write.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. 

For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]