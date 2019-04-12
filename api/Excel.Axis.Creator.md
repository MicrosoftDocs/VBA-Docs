---
title: Axis.Creator property (Excel)
keywords: vbaxl10.chm560074
f1_keywords:
- vbaxl10.chm560074
ms.prod: excel
api_name:
- Excel.Axis.Creator
ms.assetid: acbfdefc-8a21-1a64-1d7c-f3d440156d5b
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]