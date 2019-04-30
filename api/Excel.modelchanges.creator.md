---
title: ModelChanges.Creator property (Excel)
keywords: vbaxl10.chm959074
f1_keywords:
- vbaxl10.chm959074
ms.prod: excel
ms.assetid: 937eb401-ab1b-15fe-df9c-350ef13406f6
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelChanges.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ModelChanges](Excel.modelchanges.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]