---
title: WorksheetFunction.Permut method (Excel)
keywords: vbaxl10.chm137203
f1_keywords:
- vbaxl10.chm137203
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Permut
ms.assetid: dbef7a0f-bab9-83c0-9840-bb5948114b5e
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Permut method (Excel)

Returns the number of permutations for a given number of objects that can be selected from number objects. A permutation is any set or subset of objects or events where internal order is significant. Permutations are different from combinations, for which the internal order is not significant. Use this function for lottery-style probability calculations.


## Syntax

_expression_.**Permut** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - an integer that describes the number of objects.|
| _Arg2_|Required| **Double**|Number_chosen - an integer that describes the number of objects in each permutation.|

## Return value

**Double**


## Remarks

Both arguments are truncated to integers.
    
If number or number_chosen is nonnumeric, **Permut** returns the #VALUE! error value.
    
If number ≤ 0 or if number_chosen < 0, **Permut** returns the #NUM! error value.
    
If number < number_chosen, **Permut** returns the #NUM! error value.
    
The equation for the number of permutations is &nbsp; ![Formula](../images/awfpermu_ZA06051231.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]