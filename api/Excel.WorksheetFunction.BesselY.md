---
title: WorksheetFunction.BesselY method (Excel)
keywords: vbaxl10.chm137304
f1_keywords:
- vbaxl10.chm137304
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BesselY
ms.assetid: ed8e06b9-982f-b012-b6bc-ba01a6dc2fec
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.BesselY method (Excel)

Returns the Bessel function, which is also called the Weber function or the Neumann function.


## Syntax

_expression_.**BesselY** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The value at which to evaluate the function.|
| _Arg2_|Required| **Variant**|The order of the function. If n is not an integer, it is truncated.|

## Return value

**Double**


## Remarks

If x is nonnumeric, **BesselY** generates an error value.
    
If n is nonnumeric, **BesselY** generates an error value.
    
If n < 0, **BesselY** generates an error value.
    
The n-th order Bessel function of the variable x is ![Bessel function](../images/awfbsly1_ZA06051118.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]