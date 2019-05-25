---
title: WorksheetFunction.SumX2MY2 method (Excel)
keywords: vbaxl10.chm137208
f1_keywords:
- vbaxl10.chm137208
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SumX2MY2
ms.assetid: 5c1dee97-5410-2821-350e-d79a19296ef7
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.SumX2MY2 method (Excel)

Returns the sum of the difference of squares of corresponding values in two arrays.


## Syntax

_expression_.**SumX2MY2** (_Arg1_, _Arg2_)

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
    
If array_x and array_y have a different number of dimensions, **SumX2MY2** returns the #N/A error value.
    
The equation for the sum of the difference of squares is &nbsp; ![Formula](../images/awfsmx2m_ZA06051243.gif)


    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]