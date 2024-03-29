---
title: WorksheetDataConnection.Creator property (Excel)
keywords: vbaxl10.chm923074
f1_keywords:
- vbaxl10.chm923074
ms.assetid: e5965baf-48e9-be89-5cf6-76c94736d301
ms.date: 05/18/2019
ms.localizationpriority: medium
---


# WorksheetDataConnection.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[WorksheetDataConnection](Excel.worksheetdataconnection.md)** object.


## Remarks

Because the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string XCEL.


## Property value

**XLCREATOR**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]