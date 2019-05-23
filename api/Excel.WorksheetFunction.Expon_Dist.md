---
title: WorksheetFunction.Expon_Dist method (Excel)
keywords: vbaxl10.chm137365
f1_keywords:
- vbaxl10.chm137365
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Expon_Dist
ms.assetid: 19627dab-1c33-2348-389e-18a76604b237
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Expon_Dist method (Excel)

Returns the exponential distribution. Use **Expon_Dist** to model the time between events, such as how long an automated bank teller takes to deliver cash. For example, you can use **Expon_Dist** to determine the probability that the process takes at most 1 minute.


## Syntax

_expression_.**Expon_Dist** (_Arg1_, _Arg2_, _Arg3_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value of the function.|
| _Arg2_|Required| **Double**|Lambda - the parameter value.|
| _Arg3_|Required| **Boolean**|Cumulative - a logical value that indicates which form of the exponential function to provide. If cumulative is **True**, EXPONDIST returns the cumulative distribution function; if **False**, it returns the probability density function.|

## Return value

**Double**


## Remarks

If x or lambda is nonnumeric, **Expon_Dist** returns the #VALUE! error value.
    
If x < 0, **Expon_Dist** returns the #NUM! error value.
    
If lambda ≤ 0, **Expon_Dist** returns the #NUM! error value.
    
The equation for the probability density function is &nbsp; ![Formula](../images/awfxpnd1_ZA06051267.gif)
  
The equation for the cumulative distribution function is &nbsp; ![Formula](../images/awfxpnd2_ZA06051268.gif)


    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]