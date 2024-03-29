---
title: ShadowFormat.Creator property (Excel)
api_name:
- Excel.ShadowFormat.Creator
ms.assetid: 5c7397d1-dd9e-2889-ad96-6fa6510429e3
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# ShadowFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ShadowFormat](Excel.ShadowFormat.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]