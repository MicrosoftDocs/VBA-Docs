---
title: QuickAnalysis.Creator property (Excel)
keywords: vbaxl10.chm919074
f1_keywords:
- vbaxl10.chm919074
ms.prod: excel
ms.assetid: 85a25f1f-7018-9c9b-6ae4-0fd052971b70
ms.date: 05/10/2019
localization_priority: Normal
---


# QuickAnalysis.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[QuickAnalysis](Excel.quickanalysis.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]