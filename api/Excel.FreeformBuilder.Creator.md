---
title: FreeformBuilder.Creator property (Excel)
keywords: vbaxl10.chm647074
f1_keywords:
- vbaxl10.chm647074
ms.prod: excel
api_name:
- Excel.FreeformBuilder.Creator
ms.assetid: c8c85faf-83b8-1c09-b199-e711b9f3f5b4
ms.date: 04/26/2019
localization_priority: Normal
---


# FreeformBuilder.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[FreeformBuilder](Excel.FreeformBuilder.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]