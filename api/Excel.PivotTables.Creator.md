---
title: PivotTables.Creator property (Excel)
keywords: vbaxl10.chm237074
f1_keywords:
- vbaxl10.chm237074
ms.prod: excel
api_name:
- Excel.PivotTables.Creator
ms.assetid: 7af2b706-9464-765b-2653-f275ab485fe8
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotTables.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[PivotTables](Excel.PivotTables.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]