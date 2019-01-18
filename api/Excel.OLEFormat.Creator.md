---
title: OLEFormat.Creator property (Excel)
keywords: vbaxl10.chm631074
f1_keywords:
- vbaxl10.chm631074
ms.prod: excel
api_name:
- Excel.OLEFormat.Creator
ms.assetid: f7a0e432-0eda-0f6b-93da-1dcc1d9fc267
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents an [OLEFormat](Excel.OLEFormat.md) object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


[OLEFormat Object](Excel.OLEFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]