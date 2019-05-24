---
title: WorksheetFunction.Weibull method (Excel)
keywords: vbaxl10.chm137206
f1_keywords:
- vbaxl10.chm137206
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Weibull
ms.assetid: 2636d646-d867-a66b-ceba-b180e4ae69fa
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Weibull method (Excel)

Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating a device's mean time to failure.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Weibull_Dist](Excel.WorksheetFunction.Weibull_Dist.md)** method.


## Syntax

_expression_.**Weibull** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Alpha - a parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - a parameter to the distribution.|
| _Arg4_|Required| **Boolean**|Cumulative - determines the form of the function.|

## Return value

**Double**


## Remarks

If x, alpha, or beta is nonnumeric, **Weibull** returns the #VALUE! error value.
    
If x < 0, **Weibull** returns the #NUM! error value.
    
If alpha ≤ 0 or if beta ≤ 0, **Weibull** returns the #NUM! error value.
    
The equation for the Weibull cumulative distribution function is &nbsp; ![Formula](../images/awfweib1_ZA06051261.gif)

The equation for the Weibull probability density function is &nbsp; ![Formula](../images/awfweib2_ZA06051262.gif)

When alpha = 1, **Weibull** returns the exponential distribution with &nbsp; ![Formula](../images/awfweib3_ZA06051263.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]