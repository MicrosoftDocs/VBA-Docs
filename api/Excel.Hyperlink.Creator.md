---
title: Hyperlink.Creator Property (Excel)
keywords: vbaxl10.chm535074
f1_keywords:
- vbaxl10.chm535074
ms.prod: excel
api_name:
- Excel.Hyperlink.Creator
ms.assetid: f944b677-ac58-77ca-7546-2fbfc04233ae
ms.date: 06/08/2017
---


# Hyperlink.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [Hyperlink](Excel.Hyperlink.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Hyperlink Object](Excel.Hyperlink.md)

