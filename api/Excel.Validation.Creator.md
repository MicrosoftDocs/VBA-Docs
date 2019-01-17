---
title: Validation.Creator property (Excel)
keywords: vbaxl10.chm531074
f1_keywords:
- vbaxl10.chm531074
ms.prod: excel
api_name:
- Excel.Validation.Creator
ms.assetid: 664abd2c-550e-bb5e-877a-db7dc43a5c52
ms.date: 06/08/2017
localization_priority: Normal
---


# Validation.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [Validation](./Excel.Validation.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Validation Object](Excel.Validation.md)

