---
title: SortFields.Creator property (Excel)
keywords: vbaxl10.chm844074
f1_keywords:
- vbaxl10.chm844074
ms.prod: excel
api_name:
- Excel.SortFields.Creator
ms.assetid: ffdbea4a-0e5f-49cb-0345-8d6ffe8d0eef
ms.date: 05/16/2019
localization_priority: Normal
---


# SortFields.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SortFields](Excel.SortFields.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]