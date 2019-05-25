---
title: WorksheetFunction.Rept method (Excel)
keywords: vbaxl10.chm137091
f1_keywords:
- vbaxl10.chm137091
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rept
ms.assetid: acf1bf30-3722-79f3-c3ab-42c3f14aa435
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Rept method (Excel)

Repeats text a given number of times. Use **Rept** to fill a cell with a number of instances of a text string.


## Syntax

_expression_.**Rept** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text - the text that you want to repeat.|
| _Arg2_|Required| **Double**|Number_times - a positive number specifying the number of times to repeat text.|

## Return value

**String**


## Remarks

If number_times is 0 (zero), **Rept** returns "" (empty text).
    
If number_times is not an integer, it is truncated.
    
The result of the **Rept** function cannot be longer than 32,767 characters, or **Rept** returns the #VALUE! error value.
    

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]