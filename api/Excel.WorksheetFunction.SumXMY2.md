---
title: WorksheetFunction.SumXMY2 method (Excel)
keywords: vbaxl10.chm137207
f1_keywords:
- vbaxl10.chm137207
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SumXMY2
ms.assetid: 189d23bd-742e-96b7-3638-ea7af795f583
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.SumXMY2 method (Excel)

Returns the sum of squares of differences of corresponding values in two arrays.


## Syntax

_expression_.**SumXMY2** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array_x - the first array or range of values.|
| _Arg2_|Required| **Variant**|Array_y - the second array or range of values.|

## Return value

**Double**


## Remarks

The arguments should be either numbers or names, arrays, or references that contain numbers.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
If array_x and array_y have a different number of dimensions, **SumXMY2** returns the #N/A error value.
    
The equation for the sum of squared differences is &nbsp; ![Formula](../images/awfsumxm_ZA06051252.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]