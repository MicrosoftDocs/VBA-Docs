---
title: LinearGradient.Creator property (Excel)
keywords: vbaxl10.chm854074
f1_keywords:
- vbaxl10.chm854074
ms.prod: excel
api_name:
- Excel.LinearGradient.Creator
ms.assetid: 318042d1-d486-5d52-91cb-0a102ee9ae9d
ms.date: 06/08/2017
localization_priority: Normal
---


# LinearGradient.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [LinearGradient](Excel.LinearGradient.md) object.


## Return value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL. 


## See also


[LinearGradient Object](Excel.LinearGradient.md)

