---
title: WorksheetFunction.FInv method (Excel)
keywords: vbaxl10.chm137186
f1_keywords:
- vbaxl10.chm137186
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FInv
ms.assetid: 4194c2ca-a9c7-ba96-2f17-b24bcb6f4a36
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.FInv method (Excel)

Returns the inverse of the F probability distribution. If p = FDIST(x,...), then FINV(p,...) = x.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[F_Inv_RT](Excel.WorksheetFunction.F_Inv_RT.md)** and **[F_Inv](Excel.WorksheetFunction.F_Inv.md)** methods.

## Syntax

_expression_.**FInv** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability associated with the F cumulative distribution.|
| _Arg2_|Required| **Double**|Degrees_freedom1 - the numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Degrees_freedom2 - is the denominator degrees of freedom.|

## Return value

**Double**


## Remarks

The F distribution can be used in an F-test that compares the degree of variability in two data sets. For example, you can analyze income distributions in the United States and Canada to determine whether the two countries have a similar degree of income diversity.

If any argument is nonnumeric, **FInv** returns the #VALUE! error value.
    
If probability < 0 or probability > 1, **FInv** returns the #NUM! error value.
    
If degrees_freedom1 or degrees_freedom2 is not an integer, it is truncated.
    
If degrees_freedom1 < 1 or degrees_freedom1 ≥ 10^10, **FInv** returns the #NUM! error value.
    
If degrees_freedom2 < 1 or degrees_freedom2 ≥ 10^10, **FInv** returns the #NUM! error value.
    
**FInv** can be used to return critical values from the F distribution. For example, the output of an ANOVA calculation often includes data for the F statistic, F probability, and F critical value at the 0.05 significance level. To return the critical value of F, use the significance level as the probability argument to **FInv**.

Given a value for probability, **FInv** seeks that value x such that FDIST(x, degrees_freedom1, degrees_freedom2) = probability. Thus, precision of **FInv** depends on precision of **FDist**. **FInv** uses an iterative search technique. If the search has not converged after 64 iterations, the function returns the #N/A error value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]