---
title: WorksheetFunction.Small method (Excel)
keywords: vbaxl10.chm137230
f1_keywords:
- vbaxl10.chm137230
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Small
ms.assetid: d73da9a7-c518-1071-205a-042329d14918
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Small method (Excel)

Returns the k-th smallest value in a data set. Use this function to return values with a particular relative standing in a data set.


## Syntax

_expression_.**Small** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - an array or range of numerical data for which you want to determine the k-th smallest value.|
| _Arg2_|Required| **Double**|K - the position (from the smallest) in the array or range of data to return.|

## Return value

**Double**


## Remarks

If array is empty, **Small** returns the #NUM! error value.
    
If k ≤ 0 or if k exceeds the number of data points, **Small** returns the #NUM! error value.
    
If n is the number of data points in array, SMALL(array,1) equals the smallest value, and SMALL(array,n) equals the largest value.
    

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]