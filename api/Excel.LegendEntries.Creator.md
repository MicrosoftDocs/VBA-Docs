---
title: LegendEntries.Creator property (Excel)
keywords: vbaxl10.chm587074
f1_keywords:
- vbaxl10.chm587074
api_name:
- Excel.LegendEntries.Creator
ms.assetid: 9b6fe17e-a40f-7d26-bfa0-f3a5c40a1cda
ms.date: 04/27/2019
ms.localizationpriority: medium
---


# LegendEntries.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[LegendEntries](Excel.LegendEntries(object).md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]