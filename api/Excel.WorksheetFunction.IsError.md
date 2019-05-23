---
title: WorksheetFunction.IsError method (Excel)
keywords: vbaxl10.chm137076
f1_keywords:
- vbaxl10.chm137076
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IsError
ms.assetid: 87902aa7-295b-5d0b-650e-b30b8a4084c8
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.IsError method (Excel)

Checks the type of value and returns **True** or **False** depending on whether the value refers to any error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).


## Syntax

_expression_.**IsError** (_Arg1_)

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