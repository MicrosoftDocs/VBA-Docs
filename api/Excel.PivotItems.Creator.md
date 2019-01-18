---
title: PivotItems.Creator property (Excel)
keywords: vbaxl10.chm247074
f1_keywords:
- vbaxl10.chm247074
ms.prod: excel
api_name:
- Excel.PivotItems.Creator
ms.assetid: 9d055e55-5ca3-a763-cd0b-acb742f55d12
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotItems.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [PivotItems](Excel.PivotItems.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[PivotItems Object](Excel.PivotItems.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]