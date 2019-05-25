---
title: WorksheetFunction.T_Dist method (Excel)
keywords: vbaxl10.chm137383
f1_keywords:
- vbaxl10.chm137383
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Dist
ms.assetid: a6b7ad29-d00f-f779-9531-4d05bc216036
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.T_Dist method (Excel)

Returns a Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are computed.


## Syntax

_expression_.**T_Dist** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The numeric value at which to evaluate the distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - An integer that indicates the number of degrees of freedom.|
| _Arg3_|Required| **Boolean**|Cumulative - A logical value that determines the form of the function. If cumulative is **True**, **T_Dist** returns the cumulative distribution function; if **False**, it returns the probability density function.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **T_Dist** returns the #VALUE! error value.
    
If deg_freedom < 1, **T_Dist** returns the #NUM! error value.
    
If x < 0, **T_Dist** returns the #NUM! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]