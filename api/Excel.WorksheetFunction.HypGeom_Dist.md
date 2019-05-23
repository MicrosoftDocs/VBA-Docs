---
title: WorksheetFunction.HypGeom_Dist method (Excel)
keywords: vbaxl10.chm137406
f1_keywords:
- vbaxl10.chm137406
ms.prod: excel
api_name:
- Excel.WorksheetFunction.HypGeom_Dist
ms.assetid: 83fd3d7f-f9f0-fa49-863e-7ddd604b4de7
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.HypGeom_Dist method (Excel)

Returns the hypergeometric distribution. **HypGeom_Dist** returns the probability of a given number of sample successes, given the sample size, population successes, and population size. Use **HypGeom_Dist** for problems with a finite population, where each observation is either a success or a failure, and where each subset of a given size is chosen with equal likelihood.


## Syntax

_expression_.**HypGeom_Dist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Sample_s - the number of successes in the sample.|
| _Arg2_|Required| **Double**|Number_sample - the size of the sample.|
| _Arg3_|Required| **Double**|Population_s - the number of successes in the population.|
| _Arg4_|Required| **Double**|Number_population - the population size.|
| _Arg5_|Optional| **Variant**|Cumulative - a logical value that determines the form of the function. If cumulative is **True**, **HypGeom_Dist** returns the cumulative distribution function; if **False**, it returns the probability mass function.|

## Return value

**Double**


## Remarks

All arguments are truncated to integers.
    
If any argument is nonnumeric, **HypGeom_Dist** returns the #VALUE! error value.
    
If sample_s < 0 or sample_s is greater than the lesser of number_sample or population_s, **HypGeom_Dist** returns the #NUM! error value.
    
If sample_s is less than the larger of 0 or (number_sample - number_population + population_s), **HypGeom_Dist** returns the #NUM! error value.
    
If number_sample ≤ 0 or number_sample > number_population, **HypGeom_Dist** returns the #NUM! error value.
    
If population_s ≤ 0 or population_s > number_population, **HypGeom_Dist** returns the #NUM! error value.
    
If number_population ≤ 0, **HypGeom_Dist** returns the #NUM! error value.
    
The equation for the hypergeometric distribution is as follows, where:

- x = sample_s 
- n = number_sample 
- M = population_s 
- N = number_population 
   
> ![Formula](../images/awfhypge_ZA06051151.gif)

**HypGeom_Dist** is used in sampling without replacement from a finite population. 


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]