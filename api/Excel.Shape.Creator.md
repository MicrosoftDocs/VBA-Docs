---
title: Shape.Creator property (Excel)
keywords: vbaxl10.chm635074
f1_keywords:
- vbaxl10.chm635074
ms.prod: excel
api_name:
- Excel.Shape.Creator
ms.assetid: cfe75d7d-a265-5b08-35a2-58470473df39
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [Shape](./Excel.Shape.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[Shape Object](Excel.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]