---
title: SeriesLines.Creator property (Excel)
keywords: vbaxl10.chm597074
f1_keywords:
- vbaxl10.chm597074
api_name:
- Excel.SeriesLines.Creator
ms.assetid: f42923f3-78a8-5573-a707-758a39d3c301
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# SeriesLines.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SeriesLines](Excel.SeriesLines(object).md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]