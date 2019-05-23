---
title: WorksheetFunction.Covar method (Excel)
keywords: vbaxl10.chm137212
f1_keywords:
- vbaxl10.chm137212
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Covar
ms.assetid: 8e08c1c6-c4c4-9088-bd2e-3ab0edc831e2
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Covar method (Excel)

Returns covariance, the average of the products of deviations for each data point pair.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[Covariance_P](Excel.WorksheetFunction.Covar.md)** and **[Covariance_S](Excel.WorksheetFunction.Covariance_S.md)** methods.

## Syntax

_expression_.**Covar** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The first cell range of integers.|
| _Arg2_|Required| **Variant**|The second cell range of integers.|

## Return value

**Double**


## Remarks

Use covariance to determine the relationship between two data sets. For example, you can examine whether greater income accompanies greater levels of education.

The arguments must either be numbers or be names, arrays, or references that contain numbers.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
If _Arg1_ and _Arg2_ have different numbers of data points, **Covar** generates an error.
    
If either _Arg1_ or _Arg2_ is empty, **Covar** generates an error.
    
The covariance is as follows, where x and y are the sample means AVERAGE(array1) and AVERAGE(array2), and n is the sample size: 
    
> ![Formula](../images/awfcovar_ZA06051128.gif)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]