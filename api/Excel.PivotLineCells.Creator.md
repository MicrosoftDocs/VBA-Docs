---
title: PivotLineCells.Creator property (Excel)
keywords: vbaxl10.chm761074
f1_keywords:
- vbaxl10.chm761074
ms.prod: excel
api_name:
- Excel.PivotLineCells.Creator
ms.assetid: c57dddd7-1ea3-6ba6-c8de-ebc2d68f7697
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotLineCells.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[PivotLineCells](Excel.PivotLineCells.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]