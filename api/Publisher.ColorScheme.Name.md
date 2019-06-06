---
title: ColorScheme.Name property (Publisher)
keywords: vbapb10.chm2686979
f1_keywords:
- vbapb10.chm2686979
ms.prod: publisher
api_name:
- Publisher.ColorScheme.Name
ms.assetid: 8816c7d5-6dac-f1ad-f7f7-590406be5bef
ms.date: 06/06/2019
localization_priority: Normal
---


# ColorScheme.Name property (Publisher)

Returns a **String** value indicating the name of the specified object. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[ColorScheme](Publisher.ColorScheme.md)** object.


## Remarks

You can use an object's name in conjunction with the **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.

The **Name** property is the default property for the **BorderArt**, **BorderArtFormat**, and **Label** objects.


## Example

This example reports the name of the color scheme for the active publication.

```vb
MsgBox "The current color scheme is " _ 
 & ActiveDocument.ColorScheme.Name & "."
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]