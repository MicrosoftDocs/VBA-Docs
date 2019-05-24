---
title: WorksheetFunction.Quartile_Exc method (Excel)
keywords: vbaxl10.chm137377
f1_keywords:
- vbaxl10.chm137377
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quartile_Exc
ms.assetid: 2b33be15-7d3c-d8be-aae1-de100de8083c
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Quartile_Exc method (Excel)

Returns the quartile of the data set, based on percentile values from 0..1, exclusive.


## Syntax

_expression_.**Quartile_Exc** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or cell range of numeric values for which you want the quartile value.|
| _Arg2_|Required| **Double**|Quart - The value to return.|

## Return value

**Double**


## Remarks

If array is empty, **Quartile_Exc** returns the #NUM! error value.
    
If quart is not an integer, it is truncated. 
    
If quart ≤ 0 or if quart ≥ 4, **Quartile_Exc** returns the #NUM! error value.
    
**Min**, **Median**, and **Max** return the same value as **Quartile_Exc** when quart is equal to 0 (zero), 2, and 4, respectively.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]