---
title: WorksheetFunction.T_Dist_2T method (Excel)
keywords: vbaxl10.chm137384
f1_keywords:
- vbaxl10.chm137384
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Dist_2T
ms.assetid: e4927634-d94c-5bcc-7bef-ad35a315bc69
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.T_Dist_2T method (Excel)

Returns the two-tailed Student t-distribution.


## Syntax

_expression_.**T_Dist_2T** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The numeric value at which to evaluate the distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - An integer that indicates the number of degrees of freedom.|

## Return value

**Double**


## Remarks

If any argument is non-numeric, **T_Dist_2T** returns the #VALUE! error value.
    
If deg_freedom < 1, **T_Dist_2T** returns the #NUM! error value.
    
If x < 0, **T_Dist_2T** returns the #NUM! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]