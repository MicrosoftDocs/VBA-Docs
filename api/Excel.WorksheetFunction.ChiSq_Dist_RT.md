---
title: WorksheetFunction.ChiSq_Dist_RT method (Excel)
keywords: vbaxl10.chm137399
f1_keywords:
- vbaxl10.chm137399
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiSq_Dist_RT
ms.assetid: 65b8bd60-c13f-9f64-58c3-cc0ce582f939
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.ChiSq_Dist_RT method (Excel)

Returns the right-tailed probability of the chi-squared distribution.


## Syntax

_expression_.**ChiSq_Dist_RT** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The value at which you want to evaluate the distribution.|
| _Arg2_|Required| **Double**|The number of degrees of freedom.|

## Return value

**Double**


## Remarks

The χ2 distribution is associated with an χ2 test. Use the χ2 test to compare observed and expected values.

For example, a genetic experiment might hypothesize that the next generation of plants will exhibit a certain set of colors. By comparing the observed results with the expected ones, you can decide whether your original hypothesis is valid.

If either argument is nonnumeric, **ChiSq_Dist_RT** generates an error.
    
If x is negative, **ChiSq_Dist_RT** generates an error.
    
If degrees_freedom is not an integer, it is truncated.
    
If degrees_freedom < 1 or degrees_freedom > 10^10, **ChiSq_Dist_RT** generates an error.
    
**ChiSq_Dist_RT** is calculated as ChiSq_Dist_RT = P(X>x), where X is an χ2 random variable.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]