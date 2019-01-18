---
title: ModelFormatBoolean.Creator property (Excel)
keywords: vbaxl10.chm995074
f1_keywords:
- vbaxl10.chm995074
ms.assetid: b32a70e5-a6ae-e1ef-cc10-e86ca88f1578
ms.date: 06/08/2017
ms.prod: excel
localization_priority: Normal
---


# ModelFormatBoolean.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a 'ModelFormatBoolean' object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator  **property** is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[ModelFormatBoolean Object](Excel.modelformatboolean.md)


