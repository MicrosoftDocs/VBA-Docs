---
title: WorksheetFunction.NormSInv method (Excel)
keywords: vbaxl10.chm137200
f1_keywords:
- vbaxl10.chm137200
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NormSInv
ms.assetid: 88b209e4-3dc0-7c21-e175-55c1f133919e
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.NormSInv method (Excel)

Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Norm_S_Inv](Excel.WorksheetFunction.Norm_S_Inv.md)** method.


## Syntax

_expression_.**NormSInv** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability corresponding to the normal distribution.|

## Return value

**Double**


## Remarks

If probability is nonnumeric, **NormSInv** returns the #VALUE! error value.
    
If probability ≤ 0 or if probability ≥ 1, **NormSInv** returns the #NUM! error value.
    
Given a value for probability, **NormSInv** seeks that value z such that NORMSDIST(z) = probability. Thus, precision of **NormSInv** depends on precision of **NormSDist**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]