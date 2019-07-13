---
title: Master.Name property (PowerPoint)
keywords: vbapp10.chm533007
f1_keywords:
- vbapp10.chm533007
ms.prod: powerpoint
api_name:
- PowerPoint.Master.Name
ms.assetid: 1c751814-61fe-c246-d516-0d43b7757248
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.Name property (PowerPoint)

Returns or sets the name of the specified object. Read/write.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Master](PowerPoint.Master.md) object.


## Return value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, `.Shapes("Rectangle 2")` returns a reference to that shape.


## See also


[Master Object](PowerPoint.Master.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]