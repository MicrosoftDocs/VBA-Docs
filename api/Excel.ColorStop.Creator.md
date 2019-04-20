---
title: ColorStop.Creator property (Excel)
keywords: vbaxl10.chm850074
f1_keywords:
- vbaxl10.chm850074
ms.prod: excel
api_name:
- Excel.ColorStop.Creator
ms.assetid: 99789f97-d576-1be6-40c5-9cd2a5984751
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorStop.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ColorStop](Excel.ColorStop.md)** object.


## Return value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]