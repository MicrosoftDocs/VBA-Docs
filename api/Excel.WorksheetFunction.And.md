---
title: WorksheetFunction.And method (Excel)
keywords: vbaxl10.chm137092
f1_keywords:
- vbaxl10.chm137092
ms.prod: excel
api_name:
- Excel.WorksheetFunction.And
ms.assetid: 562be888-b001-5855-dfab-02cd066b1f12
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.And method (Excel)

Returns **True** if all its arguments are **True**; returns **False** if one or more arguments is **False**.


## Syntax

_expression_.**And** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|1 to 30 conditions that you want to test that can be either **True** or **False**.|

## Return value

**Boolean**


## Remarks

The arguments must evaluate to logical values such as **True** or **False**, or the arguments must be arrays or references that contain logical values.
    
If an array or reference argument contains text or empty cells, those values are ignored.
    
If the specified range contains no logical values, this method generates an error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]