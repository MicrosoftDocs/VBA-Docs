---
title: RecentFiles.Creator property (Excel)
keywords: vbaxl10.chm171074
f1_keywords:
- vbaxl10.chm171074
api_name:
- Excel.RecentFiles.Creator
ms.assetid: 83b6210e-5994-2468-f4b9-0884abc689fc
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# RecentFiles.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[RecentFiles](Excel.RecentFiles.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]