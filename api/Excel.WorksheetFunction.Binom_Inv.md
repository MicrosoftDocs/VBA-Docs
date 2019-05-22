---
title: WorksheetFunction.Binom_Inv method (Excel)
keywords: vbaxl10.chm137415
f1_keywords:
- vbaxl10.chm137415
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Binom_Inv
ms.assetid: 30af29b2-fc97-656b-d703-905caf7fcbb5
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Binom_Inv method (Excel)

Returns the inverse of the individual term binomial distribution probability.


## Syntax

_expression_.**Binom_Inv** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Trials - the number of Bernoulli trials.|
| _Arg2_|Required| **Double**|Probability_s - the probability of a success on each trial.|
| _Arg3_|Required| **Double**|Alpha - the criterion value.|

## Return value

**Double**


## Remarks

If trials, probability_s, or alpha is nonnumeric, the **Binom_Inv** method generates an error.
    
If trials is not an integer, it is truncated.
    
If trials < 0, the **Binom_Inv** method generates an error.
    
If probability_s < 0 or probability_s > 1, the **Binom_Inv** method generates an error.
    
If alpha < 0 or alpha > 1, the **Binom_Inv** method generates an error.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]