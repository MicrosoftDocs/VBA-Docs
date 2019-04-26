---
title: FormatColor.Creator property (Excel)
keywords: vbaxl10.chm801074
f1_keywords:
- vbaxl10.chm801074
ms.prod: excel
api_name:
- Excel.FormatColor.Creator
ms.assetid: 8167e66c-152d-efd7-9b8a-d98f11d4ce8c
ms.date: 04/26/2019
localization_priority: Normal
---


# FormatColor.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[FormatColor](Excel.FormatColor.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]