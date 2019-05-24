---
title: WorksheetFunction.LogInv method (Excel)
keywords: vbaxl10.chm137195
f1_keywords:
- vbaxl10.chm137195
ms.prod: excel
api_name:
- Excel.WorksheetFunction.LogInv
ms.assetid: 414a4e30-1225-279b-2981-bbb798338b18
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.LogInv method (Excel)

Use the lognormal distribution to analyze logarithmically transformed data.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality. 
> 
> For more information about the new function, see the **[LogNorm_Inv](Excel.WorksheetFunction.LogNorm_Inv.md)** method.


## Syntax

_expression_.**LogInv** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability associated with the lognormal distribution.|
| _Arg2_|Required| **Double**|Mean - the mean of ln(x).|
| _Arg3_|Required| **Double**|Standard_dev - the standard deviation of ln(x).|

## Return value

**Double**


## Remarks

If any argument is nonnumeric, **LogInv** returns the #VALUE! error value.
    
If probability ≤ 0 or probability ≥ 1, **LogInv** returns the #NUM! error value.
    
If standard_dev ≤ 0, **LogInv** returns the #NUM! error value.
    
The inverse of the lognormal distribution function is &nbsp; ![Formula](../images/awflginv_ZA06051178.gif)


    
[!include[Support and feedback](~/includes/feedback-boilerplate.md)]