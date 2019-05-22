---
title: WorksheetFunction.Even method (Excel)
keywords: vbaxl10.chm137183
f1_keywords:
- vbaxl10.chm137183
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Even
ms.assetid: f67f74fd-f3af-69d1-1b42-8295fbdb1ec3
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Even method (Excel)

Returns number rounded up to the nearest even integer. You can use this function for processing items that come in twos. For example, a packing crate accepts rows of one or two items. The crate is full when the number of items, rounded up to the nearest two, matches the crate's capacity.


## Syntax

_expression_.**Even** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the value to round.|

## Return value

**Double**


## Remarks

If number is nonnumeric, **Even** returns the #VALUE! error value.
    
Regardless of the sign of number, a value is rounded up when adjusted away from zero. If number is an even integer, no rounding occurs.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]