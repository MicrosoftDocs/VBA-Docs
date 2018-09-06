---
title: Axis.Creator Property (Excel)
keywords: vbaxl10.chm560074
f1_keywords:
- vbaxl10.chm560074
ms.prod: excel
api_name:
- Excel.Axis.Creator
ms.assetid: acbfdefc-8a21-1a64-1d7c-f3d440156d5b
ms.date: 06/08/2017
---


# Axis.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents an [Axis](Excel.Axis(Graph object).md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Axis Object](Excel.Axis(object).md)

