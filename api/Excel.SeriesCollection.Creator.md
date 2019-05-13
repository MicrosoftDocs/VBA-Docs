---
title: SeriesCollection.Creator property (Excel)
keywords: vbaxl10.chm579074
f1_keywords:
- vbaxl10.chm579074
ms.prod: excel
api_name:
- Excel.SeriesCollection.Creator
ms.assetid: 31d06934-b813-65b8-209c-f950b78ab796
ms.date: 05/14/2019
localization_priority: Normal
---


# SeriesCollection.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SeriesCollection](Excel.SeriesCollection.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]