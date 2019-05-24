---
title: WorksheetFunction.Norm_S_Inv method (Excel)
keywords: vbaxl10.chm137411
f1_keywords:
- vbaxl10.chm137411
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Norm_S_Inv
ms.assetid: 731c1354-2f2e-8fa8-3ced-576dd4d3ce1c
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Norm_S_Inv method (Excel)

Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of 0 (zero) and a standard deviation of one.


## Syntax

_expression_.**Norm_S_Inv** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability corresponding to the normal distribution.|

## Return value

**Double**


## Remarks

If probability is non-numeric, **Norm_S_Inv** returns the #VALUE! error value.
    
If probability < 0 or if probability > 1, **Norm_S_Inv** returns the #NUM! error value.
    
Given a value for probability, **Norm_S_Inv** seeks that value z such that NORM_S_DIST(z) = probability. Thus, precision of **Norm_S_Inv** depends on precision of **Norm_S_Dist**. **Norm_S_Inv** uses an iterative search technique. If the search has not converged after 100 iterations, the function returns the #N/A error value.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]