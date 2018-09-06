---
title: PivotCache.Creator Property (Excel)
keywords: vbaxl10.chm226074
f1_keywords:
- vbaxl10.chm226074
ms.prod: excel
api_name:
- Excel.PivotCache.Creator
ms.assetid: 3393e844-b6e1-f767-d993-53844536782c
ms.date: 06/08/2017
---


# PivotCache.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [PivotCache](Excel.PivotCache.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[PivotCache Object](Excel.PivotCache.md)

