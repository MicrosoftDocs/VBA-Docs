---
title: WorksheetFunction.Product method (Excel)
keywords: vbaxl10.chm137143
f1_keywords:
- vbaxl10.chm137143
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Product
ms.assetid: 8bd10224-d725-860e-dbe0-44995961df3e
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Product method (Excel)

Multiplies all the numbers given as arguments and returns the product.


## Syntax

_expression_.**Product** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 30 numbers that you want to multiply.|

## Return value

**Double**


## Remarks

Arguments that are numbers, logical values, or text representations of numbers are counted; arguments that are error values or text that cannot be translated into numbers cause errors.
    
If an argument is an array or reference, only numbers in the array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]