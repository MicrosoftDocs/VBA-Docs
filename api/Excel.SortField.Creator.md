---
title: SortField.Creator property (Excel)
keywords: vbaxl10.chm842074
f1_keywords:
- vbaxl10.chm842074
ms.prod: excel
api_name:
- Excel.SortField.Creator
ms.assetid: c9247d01-32fa-3360-7261-5287e47d6d40
ms.date: 05/16/2019
localization_priority: Normal
---


# SortField.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SortField](Excel.SortField.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]