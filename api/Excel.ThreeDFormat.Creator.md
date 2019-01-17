---
title: ThreeDFormat.Creator property (Excel)
ms.prod: excel
api_name:
- Excel.ThreeDFormat.Creator
ms.assetid: adae19fb-0ef7-6366-e70d-ff43b443419a
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [ThreeDFormat](./Excel.ThreeDFormat.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[ThreeDFormat Object](Excel.ThreeDFormat.md)

