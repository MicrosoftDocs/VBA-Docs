---
title: WorksheetFunction.BesselJ method (Excel)
keywords: vbaxl10.chm137302
f1_keywords:
- vbaxl10.chm137302
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BesselJ
ms.assetid: 9d6d4059-4c84-a79a-2143-eef4953cbf19
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.BesselJ method (Excel)

Returns the Bessel function.


## Syntax

_expression_.**BesselJ** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The value at which to evaluate the function.|
| _Arg2_|Required| **Variant**|The order of the Bessel function. If n is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If x is nonnumeric, **BesselJ** generates an error value.
    
If n is nonnumeric, **BesselJ** generates an error value.
    
If n < 0, **BesselJ** generates an error value.
    
The n-th order Bessel function of the variable x is ![Bessel function](../images/awfbslj1_ZA06051116.gif) where ![Bessel function](../images/awfbslj2_ZA06051117.gif) is the Gamma function. 
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]