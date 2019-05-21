---
title: WorksheetFunction.CountBlank method (Excel)
keywords: vbaxl10.chm137243
f1_keywords:
- vbaxl10.chm137243
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CountBlank
ms.assetid: e5446c10-ec41-ac83-5bc6-ca6ad98e3f7a
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.CountBlank method (Excel)

Counts empty cells in a specified range of cells.


## Syntax

_expression_.**CountBlank** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|The range from which you want to count the blank cells.|

## Return value

**Double**


## Remarks

Cells with formulas that return "" (empty text) are also counted. Cells with zero values are not counted.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
