---
title: WorksheetFunction.Prob method (Excel)
keywords: vbaxl10.chm137221
f1_keywords:
- vbaxl10.chm137221
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Prob
ms.assetid: 7715295d-90da-53fc-ef66-8422e829e05c
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Prob method (Excel)

Returns the probability that values in a range are between two limits. If upper_limit is not supplied, returns the probability that values in x_range are equal to lower_limit.


## Syntax

_expression_.**Prob** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|X_range - the range of numeric values of x with which there are associated probabilities.|
| _Arg2_|Required| **Variant**|Prob_range - a set of probabilities associated with values in x_range.|
| _Arg3_|Required| **Double**|Lower_limit - the lower bound on the value for which you want a probability.|
| _Arg4_|Optional| **Variant**|Upper_limit - the optional upper bound on the value for which you want a probability.|

## Return value

**Double**


## Remarks

If any value in prob_range ≤ 0 or if any value in prob_range > 1, **Prob** returns the #NUM! error value.
    
If the sum of the values in prob_range ≥ 1, **Prob** returns the #NUM! error value.
    
If upper_limit is omitted, **Prob** returns the probability of being equal to lower_limit.
    
If x_range and prob_range contain a different number of data points, **Prob** returns the #N/A error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]