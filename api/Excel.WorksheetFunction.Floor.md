---
title: WorksheetFunction.Floor method (Excel)
keywords: vbaxl10.chm137189
f1_keywords:
- vbaxl10.chm137189
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Floor
ms.assetid: c35733d5-34b9-8475-197f-4f13ae1e6c1a
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Floor method (Excel)

Rounds number down, toward zero, to the nearest multiple of significance.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[Floor_Precise](Excel.WorksheetFunction.Floor_Precise.md)** method.

## Syntax

_expression_.**Floor** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the numeric value that you want to round.|
| _Arg2_|Required| **Double**|Significance - the multiple to which you want to round.|

## Return value

**Double**


## Remarks

As long as the number and specified significance have the same sign, **Floor** rounds towards zero to the nearest multiple of significance.
    
If either argument is nonnumeric, **Floor** returns the #VALUE! error value.
    
Excel allows positive and negative multiples of significance with negative numbers. In those cases, if the significance is positive, **Floor** rounds away from zero. Otherwise, if significance is negative, **Floor** rounds towards zero.
    
For positive numbers with negative multiples of significance, Excel returns the #NUM! error value.
    
If number is an exact multiple of significance, no rounding occurs.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]