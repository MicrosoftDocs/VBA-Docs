---
title: ValueChange.Creator property (Excel)
keywords: vbaxl10.chm888074
f1_keywords:
- vbaxl10.chm888074
ms.prod: excel
api_name:
- Excel.ValueChange.Creator
ms.assetid: a1d10479-e30d-c0b2-a521-8910c5b6256e
ms.date: 06/08/2017
localization_priority: Normal
---


# ValueChange.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a '[ValueChange](Excel.ValueChange.md)' object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[ValueChange Object](Excel.ValueChange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]