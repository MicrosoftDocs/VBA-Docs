---
title: WizardValue.Name property (Publisher)
keywords: vbapb10.chm2097152
f1_keywords:
- vbapb10.chm2097152
ms.prod: publisher
api_name:
- Publisher.WizardValue.Name
ms.assetid: 51cef04a-e22f-217f-a8a4-d9931057e817
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardValue.Name property (Publisher)

Returns a **String** value indicating the name of the specified object. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[WizardValue](Publisher.WizardValue.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. 

For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]