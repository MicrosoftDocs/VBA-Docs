---
title: Model.Creator property (Excel)
keywords: vbaxl10.chm941074
f1_keywords:
- vbaxl10.chm941074
ms.prod: excel
ms.assetid: 2370b2d9-e759-8ee1-806e-fa15f59e3646
ms.date: 04/30/2019
localization_priority: Normal
---


# Model.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Model](Excel.Model.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]