---
title: WorksheetFunction.Sum Method (Excel)
keywords: vbaxl10.chm137077
f1_keywords:
- vbaxl10.chm137077
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Sum
ms.assetid: bbaf28fa-ca79-4b2d-4ace-153ca931a8c4
ms.date: 06/08/2017
---


# WorksheetFunction.Sum Method (Excel)

Adds all the numbers in a range of cells.


## Syntax

 _expression_. `Sum`( `_Arg1_` , `_Arg2_` , `_Arg3_` , `_Arg4_` , `_Arg5_` , `_Arg6_` , `_Arg7_` , `_Arg8_` , `_Arg9_` , `_Arg10_` , `_Arg11_` , `_Arg12_` , `_Arg13_` , `_Arg14_` , `_Arg15_` , `_Arg16_` , `_Arg17_` , `_Arg18_` , `_Arg19_` , `_Arg20_` , `_Arg21_` , `_Arg22_` , `_Arg23_` , `_Arg24_` , `_Arg25_` , `_Arg26_` , `_Arg27_` , `_Arg28_` , `_Arg29_` , `_Arg30_` )

 _expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 arguments for which you want the total value or sum.|

### Return value

Double


## Remarks




- Numbers, logical values, and text representations of numbers that you type directly into the list of arguments are counted. See the first and second examples following.
    
- If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, or text in the array or reference are ignored. See the third example following.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

