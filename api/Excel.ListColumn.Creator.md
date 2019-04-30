---
title: ListColumn.Creator property (Excel)
keywords: vbaxl10.chm737074
f1_keywords:
- vbaxl10.chm737074
ms.prod: excel
api_name:
- Excel.ListColumn.Creator
ms.assetid: 9dad6409-cd84-e7ef-71e0-d003ca61cdda
ms.date: 04/30/2019
localization_priority: Normal
---


# ListColumn.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ListColumn](Excel.ListColumn.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]