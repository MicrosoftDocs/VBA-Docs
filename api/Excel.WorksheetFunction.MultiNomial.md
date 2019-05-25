---
title: WorksheetFunction.MultiNomial method (Excel)
keywords: vbaxl10.chm137350
f1_keywords:
- vbaxl10.chm137350
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MultiNomial
ms.assetid: be7c63a7-a575-8139-e37e-a0431b95a07c
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.MultiNomial method (Excel)

Returns the ratio of the factorial of a sum of values to the product of factorials.


## Syntax

_expression_.**MultiNomial** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 30 values for which you want the multinomial.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **MultiNomial** returns the #VALUE! error value.
    
If any argument is less than zero, **MultiNomial** returns the #NUM! error value.
    
The multinomial is &nbsp; ![Formula](../images/awfmlnom_ZA06051208.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]