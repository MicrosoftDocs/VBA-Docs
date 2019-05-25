---
title: WorksheetFunction.Weibull_Dist method (Excel)
keywords: vbaxl10.chm137390
f1_keywords:
- vbaxl10.chm137390
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Weibull_Dist
ms.assetid: 17e5c39f-0808-2c84-a732-801fa0e342d8
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Weibull_Dist method (Excel)

Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating the mean time to failure for a device.


## Syntax

_expression_.**Weibull_Dist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Alpha - A parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - A parameter to the distribution.|
| _Arg4_|Required| **Boolean**|Cumulative - Determines the form of the function.|

## Return value

**Double**


## Remarks

If x, alpha, or beta is non-numeric, **Weibull_Dist** returns the #VALUE! error value.
    
If x < 0, **Weibull_Dist** returns the #NUM! error value.
    
If alpha ≤ 0 or if beta ≤ 0, **Weibull_Dist** returns the #NUM! error value.
    
The equation for the Weibull cumulative distribution function is &nbsp; ![Formula](../images/awfweib1_ZA06051261.gif)
  
The equation for the Weibull probability density function is &nbsp; ![Formula](../images/awfweib2_ZA06051262.gif)

When alpha = 1, **Weibull_Dist** returns the exponential distribution with &nbsp; ![Formula](../images/awfweib3_ZA06051263.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]