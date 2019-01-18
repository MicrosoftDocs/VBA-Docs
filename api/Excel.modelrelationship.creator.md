---
title: ModelRelationship.Creator property (Excel)
keywords: vbaxl10.chm937074
f1_keywords:
- vbaxl10.chm937074
ms.prod: excel
ms.assetid: 8db0510e-7e39-ba02-36d1-5190fcb9c795
ms.date: 06/08/2017
localization_priority: Normal
---


# ModelRelationship.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [ModelRelationship object (Excel)](Excel.modelrelationship.md) object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string "XCEL".


## Property value

 **XLCREATOR**


## See also



[ModelRelationship Object](Excel.modelrelationship.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]