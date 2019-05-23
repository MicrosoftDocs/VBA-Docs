---
title: WorksheetFunction.IsNA method (Excel)
keywords: vbaxl10.chm137075
f1_keywords:
- vbaxl10.chm137075
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IsNA
ms.assetid: db315cb6-cf6d-eaf6-7ffa-2b8ce18c2a09
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.IsNA method (Excel)

Checks the type of value and returns **True** or **False** depending on whether the value refers to the #N/A (value not available) error value.


## Syntax

_expression_.**IsNA** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Value - the value that you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.|

## Return value

**Boolean**


## Remarks

The value arguments of the IS functions are not converted. For example, in most other functions where a number is required, the text value 19 is converted to the number 19. However, in the formula `ISNUMBER("19")`, 19 is not converted from a text value, and the **IsNumber** function returns **False**.
    
The IS functions are useful in formulas for testing the outcome of a calculation. When combined with the IF function, they provide a method for locating errors in formulas.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

