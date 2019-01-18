---
title: DataFeedConnection.Creator property (Excel)
keywords: vbaxl10.chm927074
f1_keywords:
- vbaxl10.chm927074
ms.prod: excel
ms.assetid: 42c5d1f6-b740-dd1c-87dc-4285ad0eec08
ms.date: 06/08/2017
localization_priority: Normal
---


# DataFeedConnection.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created.  **Long** Read-only


## Syntax

_expression_. `Creator`

_expression_ A variable that represents a [DataFeedConnection object (Excel)](Excel.datafeedconnection.md) object ( [Excel](Excel(enumerations).md) ).


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string "XCEL".


## Property value

 **XLCREATOR**


## See also



[DataFeedConnection Object](Excel.datafeedconnection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]