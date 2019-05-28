---
title: WorksheetFunction.StDevP method (Excel)
keywords: vbaxl10.chm137147
f1_keywords:
- vbaxl10.chm137147
ms.prod: excel
api_name:
- Excel.WorksheetFunction.StDevP
ms.assetid: 5afcdfc9-ecaa-e63e-e162-118d8aa56042
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.StDevP method (Excel)

Calculates standard deviation based on the entire population given as arguments. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[StDev_P](Excel.WorksheetFunction.StDev_P.md)** method.


## Syntax

_expression_.**StDevP** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 30 number arguments corresponding to a population. You can also use a single array or a reference to an array instead of arguments separated by commas.|

## Return value

**Double**


## Remarks

**StDevP** assumes that its arguments are the entire population. If your data represents a sample of the population, compute the standard deviation by using **StDev**.
    
For large sample sizes, **StDev** and **StDevP** return approximately equal values.
    
The standard deviation is calculated by using the "biased" or "n" method.
    
Arguments can either be numbers or names, arrays, or references that contain numbers.
    
Logical values and text representations of numbers that you type directly into the list of arguments are counted. 
    
If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
Arguments that are error values or text that cannot be translated into numbers cause errors.
    
**StDevP** uses the following formula, where x is the sample mean AVERAGE(number1,number2,...) and n is the sample size: 
    
> ![Formula](../images/awfstdv2_ZA06051249.gif)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]