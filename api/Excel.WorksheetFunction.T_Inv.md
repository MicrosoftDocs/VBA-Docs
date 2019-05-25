---
title: WorksheetFunction.T_Inv method (Excel)
keywords: vbaxl10.chm137386
f1_keywords:
- vbaxl10.chm137386
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Inv
ms.assetid: 0104e8a3-0beb-69bb-d9b5-20c319d740f6
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.T_Inv method (Excel)

Returns the left-tailed inverse of the Student t-distribution.


## Syntax

_expression_.**T_Inv** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - The probability associated with the Student t-distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - The number of degrees of freedom with which to characterize the distribution.|

## Return value

**Double**


## Remarks

If either argument is non-numeric, **T_Inv** returns the #VALUE! error value.
    
If probability < 0 or if probability > 1, **T_Inv** returns the #NUM! error value.
    
If deg_freedom is not an integer, it is truncated.
    
If deg_freedom < 1, **T_Inv** returns the #NUM! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]