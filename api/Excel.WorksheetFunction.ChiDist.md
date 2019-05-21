---
title: WorksheetFunction.ChiDist method (Excel)
keywords: vbaxl10.chm137178
f1_keywords:
- vbaxl10.chm137178
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiDist
ms.assetid: e5d6c267-b9d6-75d9-5d6f-81b616652b74
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.ChiDist method (Excel)

Returns the one-tailed probability of the chi-squared distribution. 

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[ChiSq_Dist_RT](Excel.WorksheetFunction.ChiSq_Dist_RT.md)** and **[ChiSq_Dist](Excel.WorksheetFunction.ChiSq_Dist.md)** methods.

## Syntax

_expression_.**ChiDist** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The value at which you want to evaluate the distribution.|
| _Arg2_|Required| **Double**|The number of degrees of freedom.|

## Return value

**Double**


## Remarks

The χ2 distribution is associated with an χ2 test. Use the χ2 test to compare observed and expected values.

For example, a genetic experiment might hypothesize that the next generation of plants will exhibit a certain set of colors. By comparing the observed results with the expected ones, you can decide whether your original hypothesis is valid.

If either argument is nonnumeric, **ChiDist** generates an error.
    
If x is negative, **ChiDist** generates an error.
    
If degrees_freedom is not an integer, it is truncated.
    
If degrees_freedom < 1 or degrees_freedom > 10^10, **ChiDist** generates an error.
    
**ChiDist** is calculated as ChiDist = P(X>x), where X is an χ2 random variable.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]