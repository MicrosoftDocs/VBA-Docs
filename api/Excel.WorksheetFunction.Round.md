---
title: WorksheetFunction.Round method (Excel)
keywords: vbaxl10.chm137088
f1_keywords:
- vbaxl10.chm137088
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Round
ms.assetid: 37b1abed-ed4e-5e92-ba8d-a13f573813a0
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Round method (Excel)

Rounds a number to a specified number of digits.


## Syntax

_expression_.**Round** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the number that you want to round.|
| _Arg2_|Required| **Double**|Num_digits - specifies the number of digits to which you want to round number.|

## Return value

**Double**


## Remarks

If num_digits is greater than 0 (zero), number is rounded to the specified number of decimal places.
    
If num_digits is 0, number is rounded to the nearest integer.
    
If num_digits is less than 0, number is rounded to the left of the decimal point.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]