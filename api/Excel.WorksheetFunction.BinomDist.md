---
title: WorksheetFunction.BinomDist method (Excel)
keywords: vbaxl10.chm137177
f1_keywords:
- vbaxl10.chm137177
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BinomDist
ms.assetid: 0250970f-6a0a-ff33-8f6c-25cb632635b9
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.BinomDist method (Excel)

Returns the individual term binomial distribution probability.


## Syntax

_expression_.**BinomDist** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The number of successes in trials.|
| _Arg2_|Required| **Double**|The number of independent trials.|
| _Arg3_|Required| **Double**|The probability of success on each trial.|
| _Arg4_|Required| **Boolean**|A logical value that determines the form of the function. If cumulative is **True**, **BinomDist** returns the cumulative distribution function, which is the probability that there are at most number_s successes; if **False**, it returns the probability mass function, which is the probability that there are number_s successes.|

## Return value

**Double**


## Remarks

Use **BinomDist** in problems with a fixed number of tests or trials, when the outcomes of any trial are only success or failure, when trials are independent, and when the probability of success is constant throughout the experiment. For example, **BinomDist** can calculate the probability that two of the next three babies born are male.

Number_s and trials are truncated to integers.
    
If number_s, trials, or probability_s is nonnumeric, **BinomDist** generates an error.
    
If number_s < 0 or number_s > trials, **BinomDist** generates an error.
    
If probability_s < 0 or probability_s > 1, **BinomDist** generates an error.
    
The binomial probability mass function is ![BPM function](../images/awfbnmd1_ZA06051113.gif) where ![BPM function](../images/awfbnmd2_ZA06051114.gif) is COMBIN(n,x). 

The cumulative binomial distribution is ![BPM function](../images/awfbnmd3_ZA06051115.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]