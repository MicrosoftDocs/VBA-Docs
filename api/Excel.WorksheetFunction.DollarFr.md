---
title: WorksheetFunction.DollarFr method (Excel)
keywords: vbaxl10.chm137320
f1_keywords:
- vbaxl10.chm137320
ms.prod: excel
api_name:
- Excel.WorksheetFunction.DollarFr
ms.assetid: a024cc74-605f-7ac5-77f9-7368f8b22f8c
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.DollarFr method (Excel)

Converts a dollar price expressed as a decimal number into a dollar price expressed as a fraction. Use **DollarFr** to convert decimal numbers to fractional dollar numbers, such as securities prices.


## Syntax

_expression_.**DollarFr** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Decimal_dollar - a decimal number.|
| _Arg2_|Required| **Variant**|Fraction - the integer to use in the denominator of a fraction.|

## Return value

**Double**


## Remarks

If fraction is not an integer, it is truncated.
    
If fraction is less than 0, **DollarFr** returns the #NUM! error value.
    
If fraction is 0, **DollarFr** returns the #DIV/0! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]