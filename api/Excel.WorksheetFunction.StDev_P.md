---
title: WorksheetFunction.StDev_P method (Excel)
keywords: vbaxl10.chm137382
f1_keywords:
- vbaxl10.chm137382
ms.prod: excel
api_name:
- Excel.WorksheetFunction.StDev_P
ms.assetid: bbbfd24a-aaa9-d63d-9fa3-966949299214
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.StDev_P method (Excel)

Calculates standard deviation based on the entire population given as arguments. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).


## Syntax

_expression_.**StDev_P** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 30 number arguments corresponding to a population. You can also use a single array or a reference to an array instead of arguments separated by commas (,).|

## Return value

**Double**


## Remarks

**StDev_P** assumes that its arguments are the entire population. If your data represents a sample of the population, compute the standard deviation by using **StDev_S**.
    
For large sample sizes, **StDev_S** and **StDev_P** return approximately equal values.
    
The standard deviation is calculated by using the "biased" or "n" method.
    
Arguments can either be numbers or names, arrays, or references that contain numbers.
    
Logical values and text representations of numbers that you type directly into the list of arguments are counted. 
    
If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
Arguments that are error values or text that cannot be translated into numbers cause errors.
    
**StDev_P**  uses the following formula, where x is the sample mean AVERAGE(number1,number2,...) and n is the sample size:

> ![Formula](../images/awfstdv2_ZA06051249.gif)
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]