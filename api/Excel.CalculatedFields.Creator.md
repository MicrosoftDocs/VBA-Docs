---
title: CalculatedFields.Creator property (Excel)
keywords: vbaxl10.chm243074
f1_keywords:
- vbaxl10.chm243074
ms.prod: excel
api_name:
- Excel.CalculatedFields.Creator
ms.assetid: 95b69698-d77f-eec9-fd74-09d630ca4ae5
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedFields.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[CalculatedFields](Excel.CalculatedFields.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]