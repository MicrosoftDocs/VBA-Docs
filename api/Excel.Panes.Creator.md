---
title: Panes.Creator Property (Excel)
keywords: vbaxl10.chm357074
f1_keywords:
- vbaxl10.chm357074
ms.prod: excel
api_name:
- Excel.Panes.Creator
ms.assetid: e163a44f-7413-1e72-891e-78acc6ab2634
ms.date: 06/08/2017
---


# Panes.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [Panes](Excel.Panes.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Panes Object](Excel.Panes.md)

