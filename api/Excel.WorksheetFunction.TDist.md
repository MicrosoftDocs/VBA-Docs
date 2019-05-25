---
title: WorksheetFunction.TDist method (Excel)
keywords: vbaxl10.chm137205
f1_keywords:
- vbaxl10.chm137205
ms.prod: excel
api_name:
- Excel.WorksheetFunction.TDist
ms.assetid: fb2165bc-0643-9046-13c7-0bfbd56cde93
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.TDist method (Excel)

Returns the Percentage Points (probability) for the Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are to be computed. The t-distribution is used in the hypothesis testing of small sample data sets. Use this function in place of a table of critical values for the t-distribution.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[T_Dist_RT](Excel.WorksheetFunction.T_Dist_RT.md)**, **[T_Dist](Excel.WorksheetFunction.T_Dist.md)**, and **[T_Dist_2T](Excel.WorksheetFunction.T_Dist_2T.md)** methods.


## Syntax

_expression_.**TDist** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the numeric value at which to evaluate the distribution.|
| _Arg2_|Required| **Double**|Degrees_freedom - an integer indicating the number of degrees of freedom.|
| _Arg3_|Required| **Double**|Tails - specifies the number of distribution tails to return. If tails = 1, **TDist** returns the one-tailed distribution. If tails = 2, **TDist** returns the two-tailed distribution.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **TDist** returns the #VALUE! error value.
    
If degrees_freedom < 1, **TDist** returns the #NUM! error value.
    
The degrees_freedom and tails arguments are truncated to integers.
    
If tails is any value other than 1 or 2, **TDist** returns the #NUM! error value.
    
If x < 0, **TDist** returns the #NUM! error value.
    
If tails = 1, **TDist** is calculated as `TDIST = P(X>x)`, where X is a random variable that follows the t-distribution. 

If tails = 2, **TDist** is calculated as `TDIST = P(|X| > x) = P(X > x or X < -x)`.
    
Because x < 0 is not allowed, to use **TDist** when x < 0, note that `TDIST(-x,df,1) = 1 – TDIST(x,df,1) = P(X > -x)` and `TDIST(-x,df,2) = TDIST(x df,2) = P(|X| > x)`.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]