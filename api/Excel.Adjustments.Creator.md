---
title: Adjustments.Creator property (Excel)
ms.prod: excel
api_name:
- Excel.Adjustments.Creator
ms.assetid: 5038c1f3-8110-197b-c0f0-31c2e71bf003
ms.date: 04/04/2019
localization_priority: Normal
---


# Adjustments.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[Adjustments](Excel.Adjustments.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]