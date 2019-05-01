---
title: OLEDBErrors.Creator property (Excel)
keywords: vbaxl10.chm655074
f1_keywords:
- vbaxl10.chm655074
ms.prod: excel
api_name:
- Excel.OLEDBErrors.Creator
ms.assetid: c2143d28-5e66-5207-7d8d-82333d9de724
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBErrors.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[OLEDBErrors](Excel.OLEDBErrors.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]