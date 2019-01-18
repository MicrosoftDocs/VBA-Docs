---
title: CalloutFormat.Creator property (Excel)
ms.prod: excel
api_name:
- Excel.CalloutFormat.Creator
ms.assetid: b9c90a53-613e-7b00-401c-991f12946da5
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [CalloutFormat](Excel.CalloutFormat.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[CalloutFormat Object](Excel.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]