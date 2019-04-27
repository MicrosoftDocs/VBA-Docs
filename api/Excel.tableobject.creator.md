---
title: TableObject.Creator property (Excel)
keywords: vbaxl10.chm915074
f1_keywords:
- vbaxl10.chm915074
ms.prod: excel
ms.assetid: 978051f8-395f-a80b-b62f-ece1e78298f8
ms.date: 04/19/2019
localization_priority: Normal
---


# TableObject.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[TableObject](Excel.tableobject.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]