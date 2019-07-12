---
title: Font.Name property (PowerPoint)
keywords: vbapp10.chm575015
f1_keywords:
- vbapp10.chm575015
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Name
ms.assetid: 6798b75b-7fb8-a046-1532-a8cc41b76af8
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.Name property (PowerPoint)

Returns or sets the name of the specified object. Read/write.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Font](PowerPoint.Font.md) object.


## Return value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.


## See also


[Font Object](PowerPoint.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]