---
title: WorksheetFunction.Percentile_Exc method (Excel)
keywords: vbaxl10.chm137372
f1_keywords:
- vbaxl10.chm137372
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Percentile_Exc
ms.assetid: 56a7f7eb-c69c-0baa-c64b-68fb128c4861
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Percentile_Exc method (Excel)

Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.


## Syntax

_expression_.**Percentile_Exc** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or range of data that defines relative standing.|
| _Arg2_|Required| **Double**|K - The percentile value in the range 0..1, exclusive.|

## Return value

**Double**


## Remarks

If array is empty, **Percentile_Exc** returns the #NUM! error value.
    
If k is nonnumeric, **Percentile_Exc** returns the #VALUE! error value. 
    
If k is ≤ 0 or if k ≥ 1, **Percentile_Exc** returns the #NUM! error value. 
    
If k is not a multiple of 1/(n - 1), **Percentile_Exc** interpolates to determine the value at the k-th percentile.
    
**Percentile_Exc** will interpolate when the value for the specified percentile lies between two values in the array. If it cannot interpolate for the percentile, k specified, Excel returns a #NUM! error.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]