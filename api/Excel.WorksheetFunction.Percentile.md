---
title: WorksheetFunction.Percentile method (Excel)
keywords: vbaxl10.chm137232
f1_keywords:
- vbaxl10.chm137232
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Percentile
ms.assetid: a4918744-a7b1-28f9-4591-58c5ebf25c10
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Percentile method (Excel)

Returns the k-th percentile of values in a range. You can use this function to establish a threshold of acceptance. For example, you can decide to examine candidates who score above the 90th percentile.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[Percentile_Inc](Excel.WorksheetFunction.Percentile_Inc.md)** and **[Percentile_Exc](Excel.WorksheetFunction.Percentile_Exc.md)** methods.

## Syntax

_expression_.**Percentile** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or range of data that defines relative standing.|
| _Arg2_|Required| **Double**|K - the percentile value in the range 0..1, inclusive.|

## Return value

**Double**


## Remarks

If array is empty, **Percentile** returns the #NUM! error value.
    
If k is nonnumeric, **Percentile** returns the #VALUE! error value.
    
If k is < 0 or if k > 1, **Percentile** returns the #NUM! error value.
    
If k is not a multiple of 1/(n - 1), **Percentile** interpolates to determine the value at the k-th percentile.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]