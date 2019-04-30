---
title: ModelTable.Creator property (Excel)
keywords: vbaxl10.chm933074
f1_keywords:
- vbaxl10.chm933074
ms.prod: excel
ms.assetid: 121e3d5d-d898-1aba-2b51-aff469a7eefc
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelTable.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ModelTable](Excel.modeltable.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]