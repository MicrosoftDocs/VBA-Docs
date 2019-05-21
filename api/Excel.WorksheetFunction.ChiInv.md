---
title: WorksheetFunction.ChiInv method (Excel)
keywords: vbaxl10.chm137179
f1_keywords:
- vbaxl10.chm137179
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiInv
ms.assetid: 10b89d77-bc9f-80b0-dc31-f90c50f7e580
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.ChiInv method (Excel)

Returns the inverse of the one-tailed probability of the chi-squared distribution.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[ChiSq_Inv_RT](Excel.WorksheetFunction.ChiSq_Inv_RT.md)** and **[ChiSq_Inv](Excel.WorksheetFunction.ChiSq_Inv.md)** methods.

## Syntax

_expression_.**ChiInv** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|A probability associated with the chi-squared distribution.|
| _Arg2_|Required| **Double**|The number of degrees of freedom.|

## Return value

**Double**


## Remarks

If probability = ChiDist(x,...), then ChiInv(probability,...) = x. Use this function to compare observed results with expected ones to decide whether your original hypothesis is valid.

If either argument is nonnumeric, **ChiInv** generates an error.
    
If probability < 0 or probability > 1, **ChiInv** generates an error.
    
If degrees_freedom is not an integer, it is truncated.
    
If degrees_freedom < 1 or degrees_freedom ≥ 10^10, **ChiInv** generates an error.
    
Given a value for probability, **ChiInv** seeks that value x such that ChiDist(x, degrees_freedom) = probability. Thus, precision of **ChiInv** depends on precision of **ChiDist**. **ChiInv** uses an iterative search technique. If the search has not converged after 64 iterations, the function generates an error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]