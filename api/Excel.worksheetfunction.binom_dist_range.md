---
title: WorksheetFunction.Binom_Dist_Range method (Excel)
keywords: vbaxl10.chm137447
f1_keywords:
- vbaxl10.chm137447
ms.prod: excel
ms.assetid: 389223fe-9c1e-8aa7-8437-0ef09cbbfc3d
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Binom_Dist_Range method (Excel)

Returns the probability of a trial result using a binomial distribution.


## Syntax

_expression_.**Binom_Dist_Range**  (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|**Double**|The number of independent trials.|
| _Arg2_|Required|**Double**|The probability of success on each trial.|
| _Arg3_|Required|**Double**|The number of successes in trials.|
| _Arg4_|Optional|**Variant**|If provided, this function returns the probability that the number of successful trials shall lie between _Arg3_ and _Arg4_.|

## Return value

**Double**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]