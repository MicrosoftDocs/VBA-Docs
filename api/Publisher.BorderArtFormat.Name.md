---
title: BorderArtFormat.Name Property (Publisher)
keywords: vbapb10.chm7602179
f1_keywords:
- vbapb10.chm7602179
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.Name
ms.assetid: 742bb441-8661-b08d-8503-963421753cef
ms.date: 06/08/2017
localization_priority: Normal
---


# BorderArtFormat.Name Property (Publisher)

Returns or sets a  **String** value indicating the name of the specified object. Read/write.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a  **BorderArtFormat** object.


## Remarks

You can use an object's name in conjunction with the  **Item** method or **Item** property to return a reference to the object if the **Item** method or property for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.

The  **Name** property is the default property for the **BorderArt**,  **BorderArtFormat**, and  **Label** objects.


## See also


 [BorderArtFormat Object](Publisher.BorderArtFormat.md)

