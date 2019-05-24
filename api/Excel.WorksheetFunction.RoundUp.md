---
title: WorksheetFunction.RoundUp method (Excel)
keywords: vbaxl10.chm137157
f1_keywords:
- vbaxl10.chm137157
ms.prod: excel
api_name:
- Excel.WorksheetFunction.RoundUp
ms.assetid: daff9e6a-5ed8-b502-24c1-c4ffe01d2d0f
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.RoundUp method (Excel)

Rounds a number up, away from 0 (zero).


## Syntax

_expression_.**RoundUp** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - any real number that you want rounded up.|
| _Arg2_|Required| **Double**|Num_digits - the number of digits to which you want to round number.|

## Return value

**Double**


## Remarks

**RoundUp** behaves like **Round**, except that it always rounds a number up.
    
If num_digits is greater than 0 (zero), number is rounded up to the specified number of decimal places.
    
If num_digits is 0, number is rounded up to the nearest integer.
    
If num_digits is less than 0, number is rounded up to the left of the decimal point.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
