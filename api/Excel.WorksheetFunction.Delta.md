---
title: WorksheetFunction.Delta method (Excel)
keywords: vbaxl10.chm137295
f1_keywords:
- vbaxl10.chm137295
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Delta
ms.assetid: a8698aa3-88cf-fe5f-be57-f01daddfa4fd
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Delta method (Excel)

Tests whether two values are equal. Returns 1 if number1 = number2; otherwise, returns 0.


## Syntax

_expression_.**Delta** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number1 - the first number.|
| _Arg2_|Optional| **Variant**|Number2 - the second number. If omitted, number2 is assumed to be zero.|

## Return value

**Double**


## Remarks

Use this function to filter a set of values. For example, by summing several DELTA functions, you calculate the count of equal pairs. This function is also known as the Kronecker Delta function.

If number1 is nonnumeric, **Delta** returns the #VALUE! error value.
    
If number2 is nonnumeric, **Delta** returns the #VALUE! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]