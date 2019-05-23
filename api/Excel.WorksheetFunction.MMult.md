---
title: WorksheetFunction.MMult method (Excel)
keywords: vbaxl10.chm137139
f1_keywords:
- vbaxl10.chm137139
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MMult
ms.assetid: 8f410152-5682-2d71-007a-5fba5f884860
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.MMult method (Excel)

Returns the matrix product of two arrays. The result is an array with the same number of rows as array1 and the same number of columns as array2.

## Syntax

_expression_.**MMult** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg2_|Required| **Variant**|Array1, array2 - the arrays you want to multiply.|

## Return value

**Variant**


## Remarks

The number of columns in array1 must be the same as the number of rows in array2, and both arrays must contain only numbers. 
    
Array1 and array2 can be given as cell ranges, array constants, or references.
    
**MMult** returns the #VALUE! error when:
    
- Any cells are empty or contain text.
    
- The number of columns in array1 is different from the number of rows in array2.
    
- The size of the resulting array is equal to or greater than a total of 5,461 cells.
    
The matrix product array _a_ of two arrays _b_ and _c_ is as follows, where _i_ is the row number, and _j_ is the column number:

> ![Formula](../images/awfmmult_ZA06051209.gif)
    
Formulas that return arrays must be entered as array formulas.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]