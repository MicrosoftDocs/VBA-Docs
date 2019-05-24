---
title: WorksheetFunction.LogNorm_Dist method (Excel)
keywords: vbaxl10.chm137407
f1_keywords:
- vbaxl10.chm137407
ms.prod: excel
api_name:
- Excel.WorksheetFunction.LogNorm_Dist
ms.assetid: df3510f3-0518-9e65-f9e9-af393c3113e1
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.LogNorm_Dist method (Excel)

Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev. Use this function to analyze data that has been logarithmically transformed.


## Syntax

_expression_.**LogNorm_Dist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Mean - The mean of ln(x).|
| _Arg3_|Required| **Double**|Standard_dev - The standard deviation of ln(x).|
| _Arg4_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is **True**, **LogNorm_Dist** returns the cumulative distribution function; if **False**, it returns the probability density function.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **LogNorm_Dist** returns the #VALUE! error value.
    
If x ≤ 0 or if standard_dev ≤ 0, **LogNorm_Dist** returns the #NUM! error value.
    
The equation for the lognormal cumulative distribution function is:

> ![Equation for the lognormal cumulative distribution function](../images/LOGNORM_DIST_ZA10390996.jpg)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]