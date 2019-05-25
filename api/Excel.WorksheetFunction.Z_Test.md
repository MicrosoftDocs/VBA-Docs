---
title: WorksheetFunction.Z_Test method (Excel)
keywords: vbaxl10.chm137413
f1_keywords:
- vbaxl10.chm137413
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Z_Test
ms.assetid: 86c2af95-965f-f249-7775-65ff5c41785d
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Z_Test method (Excel)

Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, **Z_Test** returns the probability that the sample mean would be greater than the average of observations in the data set (_array_); that is, the observed sample mean.


## Syntax

_expression_.**Z_Test** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|_Array_ is the array or range of data against which to test the hypothesized population mean.|
| _Arg2_|Required| **Double**|The value to test.|
| _Arg3_|Optional| **Variant**|_Sigma_ is the population (known) standard deviation. If omitted, the sample standard deviation is used.|

## Return value

**Double**


## Remarks

If _array_ is empty, **Z_Test** returns the #N/A error value.
    
**Z_Test** is calculated as follows when _sigma_ is not omitted:

> ![The Z_TEST calculation when sigma is not omitted](../images/Z_TEST_SIGMA_ZA10391001.jpg)

**Z_Test** is calculated as follows when _sigma_ is omitted, where _x_ is the sample mean AVERAGE(_array_); _s_ is the sample standard deviation STDEV_S(_array_); and _n_ is the number of observations in the sample COUNT(_array_): 

> ![The Z_TEST calculation when sigma is omitted](../images/Z_TEST_ZA10391000.jpg)
    
**Z_Test** represents the probability that the sample mean would be greater than the observed value AVERAGE(_array_), when the underlying population mean is μ0. From the symmetry of the Normal distribution, if AVERAGE(_array_) < μ0, **Z_Test** will return a value greater than 0.5.
    
The following Excel formula can be used to calculate the two-tailed probability that the sample mean would be further from μ0 (in either direction) than AVERAGE(_array_), when the underlying population mean is μ0: 

> `=2 * MIN(Z_TEST(_array_,μ0,_sigma_), 1 - Z_TEST(_array_,μ0,_sigma_))`
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]