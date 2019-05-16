---
title: SlicerCache.Creator property (Excel)
keywords: vbaxl10.chm896074
f1_keywords:
- vbaxl10.chm896074
ms.prod: excel
api_name:
- Excel.SlicerCache.Creator
ms.assetid: 5ad84292-103d-1adb-620d-44726a3c6f0b
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerCache.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SlicerCache](Excel.SlicerCache.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]