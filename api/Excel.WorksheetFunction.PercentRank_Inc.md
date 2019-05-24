---
title: WorksheetFunction.PercentRank_Inc method (Excel)
keywords: vbaxl10.chm137375
f1_keywords:
- vbaxl10.chm137375
ms.prod: excel
api_name:
- Excel.WorksheetFunction.PercentRank_Inc
ms.assetid: 589a4d54-d7ca-84ea-2b62-dccb5e6e3ad0
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.PercentRank_Inc method (Excel)

Returns the rank of a value in a data set as a percentage (0..1, inclusive) of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use **PercentRank_Inc** to evaluate the standing of an aptitude test score among all scores for the test.


## Syntax

_expression_.**PercentRank_Inc** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or range of data with numeric values that defines relative standing.|
| _Arg2_|Required| **Double**|X - The value for which you want to know the rank.|
| _Arg3_|Optional| **Variant**|Significance - An optional value that identifies the number of significant digits for the returned percentage value. If omitted, **PercentRank_Inc** uses three digits (0.xxx).|

## Return value

**Double**


## Remarks

If array is empty, **PercentRank_Inc** returns the #NUM! error value.
    
If significance < 1, **PercentRank_Inc** returns the #NUM! error value.
    
If x does not match one of the values in array, **PercentRank_Inc** interpolates to return the correct percentage rank.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]