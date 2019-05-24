---
title: WorksheetFunction.PercentRank method (Excel)
keywords: vbaxl10.chm137233
f1_keywords:
- vbaxl10.chm137233
ms.prod: excel
api_name:
- Excel.WorksheetFunction.PercentRank
ms.assetid: c8cd2c3a-0858-27fe-b764-6bc2e7e14bf8
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.PercentRank method (Excel)

Returns the rank of a value in a data set as a percentage of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use **PercentRank** to evaluate the standing of an aptitude test score among all scores for the test.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[Percentile_Inc](Excel.WorksheetFunction.Percentile_Inc.md)** and **[Percentile_Exc](Excel.WorksheetFunction.Percentile_Exc.md)** methods.

## Syntax

_expression_.**PercentRank** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or range of data with numeric values that defines relative standing.|
| _Arg2_|Required| **Double**|X - the value for which you want to know the rank.|
| _Arg3_|Optional| **Variant**|Significance - an optional value that identifies the number of significant digits for the returned percentage value. If omitted, **PercentRank** uses three digits (0.xxx).|

## Return value

**Double**


## Remarks

If array is empty, **PercentRank** returns the #NUM! error value.
    
If significance < 1, **PercentRank** returns the #NUM! error value.
    
If x does not match one of the values in array, **PercentRank** interpolates to return the correct percentage rank.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]