---
title: WorksheetFunction.BesselK method (Excel)
keywords: vbaxl10.chm137303
f1_keywords:
- vbaxl10.chm137303
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BesselK
ms.assetid: 9b2eb52e-2b8a-3608-6410-52abccc886b3
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.BesselK method (Excel)

Returns the modified Bessel function, which is equivalent to the Bessel functions evaluated for purely imaginary arguments.


## Syntax

_expression_.**BesselK** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The value at which to evaluate the function.|
| _Arg2_|Required| **Variant**|The order of the function. If n is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If x is nonnumeric, **BesselK** generates an error value.
    
If n is nonnumeric, **BesselK** generates an error value.
    
If n < 0, **BesselK** generates an error value.
    
The n-th order modified Bessel function of the variable x is ![Bessel function](../images/awfbeslk_ZA06051112.gif) where Jn and Yn are the J and Y Bessel functions, respectively. 
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]