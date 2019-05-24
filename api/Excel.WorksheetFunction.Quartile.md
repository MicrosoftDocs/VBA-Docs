---
title: WorksheetFunction.Quartile method (Excel)
keywords: vbaxl10.chm137231
f1_keywords:
- vbaxl10.chm137231
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quartile
ms.assetid: 92893342-0ae8-a145-4b44-4236fccf2ff8
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Quartile method (Excel)

Returns the quartile of a data set. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use **Quartile** to find the top 25 percent of incomes in a population.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new functions, see the **[Quartile_Inc](Excel.WorksheetFunction.Quartile_Inc.md)** and **[Quartile_Exc](Excel.WorksheetFunction.Quartile_Exc.md)** methods.


## Syntax

_expression_.**Quartile** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or cell range of numeric values for which you want the quartile value.|
| _Arg2_|Required| **Double**|Quart - indicates which value to return.|

## Return value

**Double**


## Remarks

The following table describes the values that can be used for _Arg2_.

|If quart equals|Quartile returns|
|:-----|:-----|
|0|Minimum value|
|1|First quartile (25th percentile)|
|2|Median value (50th percentile)|
|3|Third quartile (75th percentile)|
|4|Maximum value|

If array is empty, **Quartile** returns the #NUM! error value.
    
If quart is not an integer, it is truncated.
    
If quart < 0 or if quart > 4, **Quartile** returns the #NUM! error value.
    
**Min**, **Median**, and **Max** return the same value as **Quartile** when quart is equal to 0 (zero), 2, and 4, respectively.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]