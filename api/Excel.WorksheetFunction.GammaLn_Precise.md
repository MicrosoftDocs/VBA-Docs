---
title: WorksheetFunction.GammaLn_Precise method (Excel)
keywords: vbaxl10.chm137418
f1_keywords:
- vbaxl10.chm137418
ms.prod: excel
api_name:
- Excel.WorksheetFunction.GammaLn_Precise
ms.assetid: a428c7a2-452e-575d-7d16-fd9f5023755d
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.GammaLn_Precise method (Excel)

Returns the natural logarithm of the gamma function, Γ(x).


## Syntax

_expression_.**GammaLn_Precise** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value for which you want to calculate **GammaLn**.|

## Return value

**Double**


## Remarks

If x is nonnumeric, the **GammaLn_Precise** method generates an error.
    
If x ≤ 0, the **GammaLn_Precise** method generates an error.
    
The number e raised to the GAMMALN(i) power, where i is an integer, returns the same result as (i - 1)!.
    
**GammaLn** is calculated as follows:

> ![Formula](../images/awfgamm1_ZA06051143.gif) &nbsp; where &nbsp; ![Formula](../images/awfgamm2_ZA06051144.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]