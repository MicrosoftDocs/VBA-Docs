---
title: WorksheetFunction.Gcd method (Excel)
keywords: vbaxl10.chm137349
f1_keywords:
- vbaxl10.chm137349
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Gcd
ms.assetid: 243cc3ae-d35d-66a1-2db5-d5542dec548e
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.Gcd method (Excel)

Returns the greatest common divisor of two or more integers. The greatest common divisor is the largest integer that divides both number1 and number2 without a remainder.


## Syntax

_expression_.**Gcd** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 29 values. If any value is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **Gcd** returns the #VALUE! error value.
    
If any argument is less than zero, **Gcd** returns the #NUM! error value.
    
One divides any value evenly.
    
A prime number has only itself and one as even divisors.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]