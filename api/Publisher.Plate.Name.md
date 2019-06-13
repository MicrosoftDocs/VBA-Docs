---
title: Plate.Name property (Publisher)
keywords: vbapb10.chm2883591
f1_keywords:
- vbapb10.chm2883591
ms.prod: publisher
api_name:
- Publisher.Plate.Name
ms.assetid: 47453f6b-2f5b-17ba-5787-4701636ccd72
ms.date: 06/13/2019
localization_priority: Normal
---


# Plate.Name property (Publisher)

Returns a **String** value indicating the name of the specified object. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Plate](Publisher.Plate.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. 

For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]