---
title: WorksheetFunction.NormInv method (Excel)
keywords: vbaxl10.chm137199
f1_keywords:
- vbaxl10.chm137199
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NormInv
ms.assetid: dfc745a0-6433-bb63-324f-1d22447406bd
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.NormInv method (Excel)

Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Norm_Inv](Excel.WorksheetFunction.Norm_Inv.md)** method.

## Syntax

_expression_.**NormInv** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability corresponding to the normal distribution.|
| _Arg2_|Required| **Double**|Mean - the arithmetic mean of the distribution.|
| _Arg3_|Required| **Double**|Standard_dev - the standard deviation of the distribution.|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **NormInv** returns the #VALUE! error value.
    
If probability < 0 or if probability > 1, **NormInv** returns the #NUM! error value.
    
If standard_dev ≤ 0, **NormInv** returns the #NUM! error value.
    
If mean = 0 and standard_dev = 1, **NormInv** uses the standard normal distribution (see **NormSInv**).
    
Given a value for probability, **NormInv** seeks that value x such that NORMDIST(x, mean, standard_dev, TRUE) = probability. Thus, precision of **NormInv** depends on precision of **NormDist**. **NormInv** uses an iterative search technique. If the search has not converged after 100 iterations, the function returns the #N/A error value.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]