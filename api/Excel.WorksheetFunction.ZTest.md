---
title: WorksheetFunction.ZTest method (Excel)
keywords: vbaxl10.chm137228
f1_keywords:
- vbaxl10.chm137228
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ZTest
ms.assetid: 24d85668-2502-14b5-73b7-24a5dae7c332
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.ZTest method (Excel)

Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, **ZTest** returns the probability that the sample mean would be greater than the average of observations in the data set (_array_); that is, the observed sample mean.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Z_Test](Excel.WorksheetFunction.Z_Test.md)** method.


## Syntax

_expression_.**ZTest** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|_Array_ is the array or range of data against which to test the hypothesized population mean.|
| _Arg2_|Required| **Double**| The value to test.|
| _Arg3_|Optional| **Variant**|_Sigma_ is the population (known) standard deviation. If omitted, the sample standard deviation is used.|

## Return value

**Double**


## Remarks

If _array_ is empty, **ZTest** returns the #N/A error value.
    
**ZTest** is calculated as follows when _sigma_ is not omitted:

> ![Formula](../images/awfztest_ZA06051270.gif) 

**ZTest** is calculated as follows when _sigma_ is omitted, where _x_ is the sample mean AVERAGE(_array_), _s_ is the sample standard deviation STDEV(_array_), and _n_ is the number of observations in the sample COUNT(_array_): 

> ![Formula](../images/awfztsta_ZA06054798.gif)
    
**ZTest** represents the probability that the sample mean would be greater than the observed value AVERAGE(_array_), when the underlying population mean is μ0. From the symmetry of the Normal distribution, if AVERAGE(_array_) < μ0, **ZTest** will return a value greater than 0.5.
    
The following Excel formula can be used to calculate the two-tailed probability that the sample mean would be further from μ0 (in either direction) than AVERAGE(_array_), when the underlying population mean is μ0: 

> `=2 * MIN(ZTEST(_array_,μ0,_sigma_), 1 - ZTEST(_array_,μ0,_sigma_))`.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]