---
title: ModelTableColumns.Creator property (Excel)
keywords: vbaxl10.chm931074
f1_keywords:
- vbaxl10.chm931074
ms.prod: excel
ms.assetid: 7aaccf6c-547e-0414-5722-22cdb1b833d1
ms.date: 06/08/2017
localization_priority: Normal
---


# ModelTableColumns.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ModelTableColumns](Excel.modeltablecolumns.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

 **XLCREATOR**


## See also



[ModelTableColumns Object](Excel.modeltablecolumns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]