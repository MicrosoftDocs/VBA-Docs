---
title: PivotField.Creator property (Excel)
keywords: vbaxl10.chm239074
f1_keywords:
- vbaxl10.chm239074
ms.prod: excel
api_name:
- Excel.PivotField.Creator
ms.assetid: 4e80362c-b014-95bd-ef8e-705a5f58cc57
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]