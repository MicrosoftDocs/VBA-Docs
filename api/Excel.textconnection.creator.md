---
title: TextConnection.Creator property (Excel)
keywords: vbaxl10.chm925074
f1_keywords:
- vbaxl10.chm925074
ms.prod: excel
ms.assetid: 64293b6f-41c7-54a5-9fcb-f4d19d60b0e6
ms.date: 05/17/2019
localization_priority: Normal
---


# TextConnection.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[TextConnection](Excel.TextConnection.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]