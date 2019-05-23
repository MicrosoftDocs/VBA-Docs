---
title: WorksheetFunction.ImDiv method (Excel)
keywords: vbaxl10.chm137274
f1_keywords:
- vbaxl10.chm137274
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImDiv
ms.assetid: 6379d38c-032c-da1e-b71d-cb32f59df51d
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ImDiv method (Excel)

Returns the quotient of two complex numbers in x + yi or x + yj text format.


## Syntax

_expression_.**ImDiv** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber1 - the complex numerator or dividend.|
| _Arg2_|Required| **Variant**|Inumber2 - the complex denominator or divisor.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
The quotient of two complex numbers is &nbsp; ![Formula](../images/awfimdiv_ZA06051158.gif)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]