---
title: WorksheetFunction.Norm_S_Dist method (Excel)
keywords: vbaxl10.chm137410
f1_keywords:
- vbaxl10.chm137410
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Norm_S_Dist
ms.assetid: ea17ac4a-82dc-ce24-0b3f-dc0452d805c6
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Norm_S_Dist method (Excel)

Returns the standard normal cumulative distribution function. The distribution has a mean of 0 (zero) and a standard deviation of one. Use this function in place of a table of standard normal curve areas.


## Syntax

_expression_.**Norm_S_Dist** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Z - The value for which you want the distribution.|
| _Arg2_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is **True**, **Norm_S_Dist** returns the cumulative distribution function; if **False**, it returns the probability mass function.|

## Return value

**Double**


## Remarks

If z is non-numeric, **Norm_S_Dist** returns the #VALUE! error value.
    
The equation for the standard normal cumulative distribution function is:
    
> ![Equation](../images/abbf5ae3-a27b-4e9c-eff8-009885a4ccf2.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]