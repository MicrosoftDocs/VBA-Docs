---
title: ModelColumnName.Creator property (Excel)
keywords: vbaxl10.chm961074
f1_keywords:
- vbaxl10.chm961074
ms.prod: excel
ms.assetid: ea92c791-ff11-9137-e354-9e3e84993932
ms.date: 06/08/2017
localization_priority: Normal
---


# ModelColumnName.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [ModelColumnName object (Excel)](Excel.modelcolumnname.md) object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string "XCEL".


## Property value

 **XLCREATOR**


## See also



[ModelColumnName Object](Excel.modelcolumnname.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]