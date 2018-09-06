---
title: Series.Creator Property (Excel)
keywords: vbaxl10.chm577074
f1_keywords:
- vbaxl10.chm577074
ms.prod: excel
api_name:
- Excel.Series.Creator
ms.assetid: f0c855a2-6901-be4f-13e2-426b97d34ef8
ms.date: 06/08/2017
---


# Series.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [Series](Excel.Series(Graph object).md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Series Object](Excel.Series(object).md)

