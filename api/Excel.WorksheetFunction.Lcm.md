---
title: WorksheetFunction.Lcm method (Excel)
keywords: vbaxl10.chm137351
f1_keywords:
- vbaxl10.chm137351
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Lcm
ms.assetid: 42092d1d-1328-5c05-298c-3b9a77a5a0bd
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Lcm method (Excel)

Returns the least common multiple of integers. The least common multiple is the smallest positive integer that is a multiple of all integer arguments number1, number2, and so on. Use **Lcm** to add fractions with different denominators.


## Syntax

_expression_.**Lcm** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number1, number2... - 1 to 29 values for which you want the least common multiple. If value is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **Lcm** returns the #VALUE! error value.
    
If any argument is less than zero, **Lcm** returns the #NUM! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]