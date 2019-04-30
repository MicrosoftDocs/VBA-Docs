---
title: ModelFormatDecimalNumber.Creator property (Excel)
keywords: vbaxl10.chm985074
f1_keywords:
- vbaxl10.chm985074
ms.assetid: 106db87c-9b52-1e74-e899-3da9de73bd3e
ms.date: 05/01/2019
ms.prod: excel
localization_priority: Normal
---


# ModelFormatDecimalNumber.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ModelFormatDecimalNumber](Excel.modelformatdecimalnumber.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]