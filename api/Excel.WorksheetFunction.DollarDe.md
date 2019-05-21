---
title: WorksheetFunction.DollarDe method (Excel)
keywords: vbaxl10.chm137319
f1_keywords:
- vbaxl10.chm137319
ms.prod: excel
api_name:
- Excel.WorksheetFunction.DollarDe
ms.assetid: 626462e2-3415-1552-eb7e-8f7bb5346852
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.DollarDe method (Excel)

Converts a dollar price expressed as a fraction into a dollar price expressed as a decimal number. Use **DollarDe** to convert fractional dollar numbers, such as securities prices, to decimal numbers.


## Syntax

_expression_.**DollarDe** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Fractional_dollar - a number expressed as a fraction.|
| _Arg2_|Required| **Variant**|Fraction - the integer to use in the denominator of the fraction.|

## Return value

**Double**


## Remarks

If fraction is not an integer, it is truncated.
    
If fraction is less than 0, **DollarDe** returns the #NUM! error value.
    
If fraction is 0, **DollarDe** returns the #DIV/0! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]