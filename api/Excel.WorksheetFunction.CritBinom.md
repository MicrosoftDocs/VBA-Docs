---
title: WorksheetFunction.CritBinom method (Excel)
keywords: vbaxl10.chm137182
f1_keywords:
- vbaxl10.chm137182
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CritBinom
ms.assetid: df9bb77f-b3b5-3e2b-d0b1-f42aabe9c14a
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.CritBinom method (Excel)

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.


## Syntax

_expression_.**CritBinom** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The number of Bernoulli trials.|
| _Arg2_|Required| **Double**|The probability of a success on each trial.|
| _Arg3_|Required| **Double**|The criterion value.|

## Return value

**Double**


## Remarks

Use this function for quality assurance applications. For example, use **CritBinom** to determine the greatest number of defective parts that are allowed to come off an assembly line run without rejecting the entire lot.

If any argument is nonnumeric, **CritBinom** generates an error.
    
If trials is not an integer, it is truncated.
    
If trials < 0, **CritBinom** generates an error.
    
If probability_s is < 0 or probability_s > 1, **CritBinom** generates an error.
    
If alpha < 0 or alpha > 1, **CritBinom** generates an error.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]