---
title: WorksheetFunction.ImSum method (Excel)
keywords: vbaxl10.chm137289
f1_keywords:
- vbaxl10.chm137289
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSum
ms.assetid: 154d2034-8933-7b20-2cae-92580ada7250
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ImSum method (Excel)

Returns the sum of two or more complex numbers in x + yi or x + yj text format.


## Syntax

_expression_.**ImSum** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_, _Arg8_, _Arg9_, _Arg10_, _Arg11_, _Arg12_, _Arg13_, _Arg14_, _Arg15_, _Arg16_, _Arg17_, _Arg18_, _Arg19_, _Arg20_, _Arg21_, _Arg22_, _Arg23_, _Arg24_, _Arg25_, _Arg26_, _Arg27_, _Arg28_, _Arg29_, _Arg30_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Inumber1, inumber2, ... - 1 to 29 complex numbers to add.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
The sum of two complex numbers is &nbsp; ![Formula](../images/awfimsum_ZA06051172.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]