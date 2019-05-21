---
title: WorksheetFunction.BetaDist method (Excel)
keywords: vbaxl10.chm137174
f1_keywords:
- vbaxl10.chm137174
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BetaDist
ms.assetid: 0408bf55-6bfb-7b73-34e2-c1fd2a1b93c9
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.BetaDist method (Excel)

Returns the beta cumulative distribution function.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Beta_Dist](Excel.WorksheetFunction.Beta_Dist.md)** method.

## Syntax

_expression_.**BetaDist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The value between A and B at which to evaluate the function.|
| _Arg2_|Required| **Double**|A parameter of the distribution.|
| _Arg3_|Required| **Double**|A parameter of the distribution.|
| _Arg4_|Optional| **Variant**|An optional lower bound to the interval of x.|
| _Arg5_|Optional| **Variant**|An optional upper bound to the interval of x.|

## Return value

**Double**


## Remarks

The beta distribution is commonly used to study variation in the percentage of something across samples, such as the fraction of the day people spend watching television.

If any argument is nonnumeric, **BetaDist** returns the #VALUE! error value.
    
If alpha ≤ 0 or beta ≤ 0, **BetaDist** generates an error value.
    
If x < A, x > B, or A = B, **BetaDist** generates an error value.
    
If you omit values for A and B, **BetaDist** uses the standard cumulative beta distribution, so that A = 0 and B = 1.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]