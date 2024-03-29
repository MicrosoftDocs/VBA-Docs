---
title: WorksheetFunction.GeoMean method (Excel)
keywords: vbaxl10.chm137223
f1_keywords:
- vbaxl10.chm137223
api_name:
- Excel.WorksheetFunction.GeoMean
ms.assetid: 9b999236-a220-10b2-890e-8d863c74bc74
ms.date: 05/23/2019
ms.localizationpriority: medium
---


# WorksheetFunction.GeoMean method (Excel)

Returns the geometric mean of an array or range of positive data. For example, you can use **GeoMean** to calculate average growth rate given compound interest with variable rates.


## Syntax

_expression_.**GeoMean** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2... - 1 to 30 arguments for which you want to calculate the mean. You can also use a single array or a reference to an array instead of arguments separated by commas.|

## Return value

**Double**


## Remarks

Arguments can either be numbers or names, arrays, or references that contain numbers.
    
Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
Arguments that are error values or text that cannot be translated into numbers cause errors.
    
If any data point ≤ 0, **GeoMean** returns the #NUM! error value.
    
The equation for the geometric mean is:

> ![Formula](../images/awfgeomn_ZA06051145.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]