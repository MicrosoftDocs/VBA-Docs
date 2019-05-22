---
title: WorksheetFunction.ImSqrt method (Excel)
keywords: vbaxl10.chm137277
f1_keywords:
- vbaxl10.chm137277
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSqrt
ms.assetid: 095ecba9-c987-8b58-f07e-d0f79436d650
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ImSqrt method (Excel)

Returns the square root of a complex number in x + yi or x + yj text format.


## Syntax

_expression_.**ImSqrt** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the square root.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
The square root of a complex number is:

> ![Formula](../images/awfimsq1_ZA06051168.gif)

where: 

> ![Formula](../images/awfimsq2_ZA06051169.gif)

and: 

> ![Formula](../images/awfimsq3_ZA06051170.gif)

and: 

> ![Formula](../images/awfimar3_ZA06051155.gif)


    
[!include[Support and feedback](~/includes/feedback-boilerplate.md)]