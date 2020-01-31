---
title: WorksheetFunction.Asc method (Excel)
keywords: vbaxl10.chm137246
f1_keywords:
- vbaxl10.chm137246
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Asc
ms.assetid: c89ee3d9-1a3b-6a85-7e5e-b8c3049d63a0
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Asc method (Excel)

For double-byte character set (DBCS) languages, changes full-width (double-byte) characters to half-width (single-byte) characters.


## Syntax

_expression_.**Asc** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|The text or a reference to a cell that contains the text that you want to change. If the text does not contain any full-width letters, the text is not changed.|

## Return value

**String**

## Remarks

**WorksheetFunction.Asc()** is not equivalent to **[Asc()](../../language/reference/user-interface-help/asc-function)**, which converts a string to an integer.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
