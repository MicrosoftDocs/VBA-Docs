---
title: WorksheetFunction.BesselI method (Excel)
keywords: vbaxl10.chm137305
f1_keywords:
- vbaxl10.chm137305
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BesselI
ms.assetid: 06bce6ff-a7cb-d8c7-2d80-d9fd54f9324b
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.BesselI method (Excel)

Returns the modified Bessel function, which is equivalent to the Bessel function evaluated for purely imaginary arguments.


## Syntax

_expression_.**BesselI** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The value at which to evaluate the function.|
| _Arg2_|Required| **Variant**|The order of the Bessel function. If n is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If x is nonnumeric, **BesselI** returns the #VALUE! error value.
    
If n is nonnumeric, **BesselI** generates an error value.
    
If n < 0, **BesselI** generates an error value.
    
The n-th order modified Bessel function of the variable x is ![Bessel function](../images/awfbesli_ZA06051111.gif)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]