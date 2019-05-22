---
title: WorksheetFunction.IfError method (Excel)
keywords: vbaxl10.chm137357
f1_keywords:
- vbaxl10.chm137357
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IfError
ms.assetid: 864812c0-990e-2e99-3c3b-05fe5210cf16
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.IfError method (Excel)

Returns a value that you specify if a formula evaluates to an error; otherwise, returns the result of the formula. Use the **IfError** function to trap and handle errors in a formula.


## Syntax

_expression_.**IfError** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Value - the argument that is checked for an error.|
| _Arg2_|Required| **Variant**|Value_if_error - the value to return if the formula evaluates to an error. The following error types are evaluated: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!.|

## Return value

**Variant**


## Remarks

If value or value_if_error is an empty cell, **IfError** treats it as an empty string value ("").
    
If value is an array formula, **IfError** returns an array of results for each cell in the range specified in value. 



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
