---
title: WorksheetFunction.GammaDist method (Excel)
keywords: vbaxl10.chm137190
f1_keywords:
- vbaxl10.chm137190
ms.prod: excel
api_name:
- Excel.WorksheetFunction.GammaDist
ms.assetid: fa290089-e6e0-4354-f28c-49f1a702dca5
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.GammaDist method (Excel)

Returns the gamma distribution. You can use this function to study variables that may have a skewed distribution. The gamma distribution is commonly used in queuing analysis.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Gamma_Dist](Excel.WorksheetFunction.Gamma_Dist.md)** method.

## Syntax

_expression_.**GammaDist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value at which you want to evaluate the distribution.|
| _Arg2_|Required| **Double**|Alpha - a parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - a parameter to the distribution. If beta = 1, **GammaDist** returns the standard gamma distribution.|
| _Arg4_|Required| **Boolean**|Cumulative - a logical value that determines the form of the function. If cumulative is **True**, **GammaDist** returns the cumulative distribution function; if **False**, it returns the probability density function.|

## Return value

**Double**


## Remarks

If x, alpha, or beta is nonnumeric, **GammaDist** returns the #VALUE! error value.
    
If x < 0, **GammaDist** returns the #NUM! error value.
    
If alpha ≤ 0 or if beta ≤ 0, **GammaDist** returns the #NUM! error value.
    
The equation for the gamma probability density function is:

> ![Formula](../images/awfgmdi1_ZA06051146.gif)

The standard gamma probability density function is:

> ![Formula](../images/awfgmdi2_ZA06051147.gif)

When alpha = 1, **GammaDist** returns the exponential distribution with:

> ![Formula](../images/awfgmdi3_ZA06051148.gif)

For a positive integer n, when alpha = n/2, beta = 2, and cumulative = **True**, **GammaDist** returns (1 - CHIDIST(x)) with n degrees of freedom.
    
When alpha is a positive integer, **GammaDist** is also known as the Erlang distribution.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]