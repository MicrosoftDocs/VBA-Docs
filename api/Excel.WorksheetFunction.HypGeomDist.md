---
title: WorksheetFunction.HypGeomDist method (Excel)
keywords: vbaxl10.chm137193
f1_keywords:
- vbaxl10.chm137193
ms.prod: excel
api_name:
- Excel.WorksheetFunction.HypGeomDist
ms.assetid: 93d92614-a731-2390-ea8e-bb440e7188da
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.HypGeomDist method (Excel)

Returns the hypergeometric distribution. HYPGEOMDIST returns the probability of a given number of sample successes, given the sample size, population successes, and population size. Use HYPGEOMDIST for problems with a finite population, where each observation is either a success or a failure, and where each subset of a given size is chosen with equal likelihood.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [HypGeom_Dist](Excel.WorksheetFunction.HypGeom_Dist.md) method.

## Syntax

_expression_. `HypGeomDist`( `_Arg1_` , `_Arg2_` , `_Arg3_` , `_Arg4_` )

_expression_ A variable that represents a '[WorksheetFunction](Excel.WorksheetFunction.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Sample_s - the number of successes in the sample.|
| _Arg2_|Required| **Double**|Number_sample - the size of the sample.|
| _Arg3_|Required| **Double**|Population_s - the number of successes in the population.|
| _Arg4_|Required| **Double**|Number_population - the population size.|

## Return value

Double


## Remarks




- All arguments are truncated to integers.
    
- If any argument is nonnumeric, HYPGEOMDIST returns the #VALUE! error value.
    
- If sample_s < 0 or sample_s is greater than the lesser of number_sample or population_s, HYPGEOMDIST returns the #NUM! error value.
    
- If sample_s is less than the larger of 0 or (number_sample - number_population + population_s), HYPGEOMDIST returns the #NUM! error value.
    
- If number_sample ? 0 or number_sample > number_population, HYPGEOMDIST returns the #NUM! error value.
    
- If population_s ? 0 or population_s > number_population, HYPGEOMDIST returns the #NUM! error value.
    
- If number_population ? 0, HYPGEOMDIST returns the #NUM! error value.
    
- The equation for the hypergeometric distribution is:
![Formula](../images/awfhypge_ZA06051151.gif)where: x = sample_s n = number_sample M = population_s N = number_population HYPGEOMDIST is used in sampling without replacement from a finite population. 
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

