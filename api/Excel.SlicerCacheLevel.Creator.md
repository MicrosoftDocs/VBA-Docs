---
title: SlicerCacheLevel.Creator property (Excel)
keywords: vbaxl10.chm900074
f1_keywords:
- vbaxl10.chm900074
api_name:
- Excel.SlicerCacheLevel.Creator
ms.assetid: 9d590acb-150d-3573-534d-436778fdc61b
ms.date: 05/16/2019
ms.localizationpriority: medium
---


# SlicerCacheLevel.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SlicerCacheLevel](Excel.SlicerCacheLevel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]