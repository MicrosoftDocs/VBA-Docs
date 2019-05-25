---
title: WorksheetFunction.Var method (Excel)
keywords: vbaxl10.chm137100
f1_keywords:
- vbaxl10.chm137100
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Var
ms.assetid: 8e6871ad-ed1e-cc64-3bf1-5470c41cbb96
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Var method (Excel)

Estimates variance based on a sample.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Var_S](Excel.WorksheetFunction.Var_S.md)** method.


## Syntax

_expression_.**Var** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 30 number arguments corresponding to a sample of a population.|

## Return value

**Double**


## Remarks

**Var** assumes that its arguments are a sample of the population. If your data represents the entire population, compute the variance by using **VarP**.
    
Arguments can either be numbers or names, arrays, or references that contain numbers. 
    
Logical values and text representations of numbers that you type directly into the list of arguments are counted. 
    
If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
Arguments that are error values or text that cannot be translated into numbers cause errors.
    
**Var** uses the following formula, where x is the sample mean AVERAGE(number1,number2,...) and n is the sample size: 
    
> ![Formula](../images/awfvarp_ZA06051259.gif)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]