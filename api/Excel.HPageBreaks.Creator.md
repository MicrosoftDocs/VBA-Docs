---
title: HPageBreaks.Creator Property (Excel)
keywords: vbaxl10.chm163074
f1_keywords:
- vbaxl10.chm163074
ms.prod: excel
api_name:
- Excel.HPageBreaks.Creator
ms.assetid: 9f783dd5-fd32-2360-642b-40c781f48cbe
ms.date: 06/08/2017
---


# HPageBreaks.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [HPageBreaks](Excel.HPageBreaks.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[HPageBreaks Object](Excel.HPageBreaks.md)

