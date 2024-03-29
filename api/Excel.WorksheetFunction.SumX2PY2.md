---
title: WorksheetFunction.SumX2PY2 method (Excel)
keywords: vbaxl10.chm137209
f1_keywords:
- vbaxl10.chm137209
api_name:
- Excel.WorksheetFunction.SumX2PY2
ms.assetid: 9767cc52-2f94-c57d-2410-1c3081a6b6e4
ms.date: 05/25/2019
ms.localizationpriority: medium
---


# WorksheetFunction.SumX2PY2 method (Excel)

Returns the sum of the sum of squares of corresponding values in two arrays. The sum of the sum of squares is a common term in many statistical calculations.


## Syntax

_expression_.**SumX2PY2** (_Arg1_, _Arg2_)

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
    
If array_x and array_y have a different number of dimensions, **SumX2PY2** returns the #N/A error value.
    
The equation for the sum of the sum of squares is &nbsp; ![Formula](../images/awfsmx2p_ZA06051244.gif)


    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]