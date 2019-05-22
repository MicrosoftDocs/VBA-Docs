---
title: WorksheetFunction.F_Dist method (Excel)
keywords: vbaxl10.chm137402
f1_keywords:
- vbaxl10.chm137402
ms.prod: excel
api_name:
- Excel.WorksheetFunction.F_Dist
ms.assetid: 7b18fd63-120f-fddf-a20a-00d4182778a5
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.F_Dist method (Excel)

Returns the F probability distribution.


## Syntax

_expression_.**F_Dist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Deg_freedom1 - The numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Deg_freedom2 - The denominator degrees of freedom.|
| _Arg4_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is **True**, F_DIST returns the cumulative distribution function; if **False**, it returns the probability density function.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **F_Dist** returns the #VALUE! error value.
    
If x is negative, **F_Dist** returns the #NUM! error value.
    
If deg_freedom1 or deg_freedom2 is not an integer, it is truncated.
    
If deg_freedom1 < 1, **F_Dist** returns the #NUM! error value.
    
If deg_freedom < 1, **F_Dist** returns the #NUM! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]