---
title: WorksheetFunction.Erf method (Excel)
keywords: vbaxl10.chm137300
f1_keywords:
- vbaxl10.chm137300
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Erf
ms.assetid: 1c40c49d-6866-084e-7b35-4caf3d97971e
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Erf method (Excel)

Returns the error function integrated between lower_limit and upper_limit.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Erf_Precise](Excel.WorksheetFunction.Erf_Precise.md)** method.

## Syntax

_expression_.**Erf** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Lower_limit - the lower bound for integrating **Erf**.|
| _Arg2_|Optional| **Variant**|Upper_limit - the upper bound for integrating **Erf**. If omitted, **Erf** integrates between zero and lower_limit.|

## Return value

**Double**


## Remarks

If lower_limit is nonnumeric, **Erf** returns the #VALUE! error value.
    
If lower_limit is negative, **Erf** returns the #NUM! error value.
    
If upper_limit is nonnumeric, **Erf** returns the #VALUE! error value.
    
If upper_limit is negative, **Erf** returns the #NUM! error value.

![Formula](../images/awferf1_ZA06051136.gif)

<br/>

![Formula](../images/awferf2_ZA06051137.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]