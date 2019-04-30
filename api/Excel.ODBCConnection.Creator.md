---
title: ODBCConnection.Creator property (Excel)
keywords: vbaxl10.chm795074
f1_keywords:
- vbaxl10.chm795074
ms.prod: excel
api_name:
- Excel.ODBCConnection.Creator
ms.assetid: 4af01c0a-df29-22fb-d5f9-ccbe2f6ab929
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCConnection.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[ODBCConnection](Excel.ODBCConnection.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]