---
title: WorksheetFunction.ImSub method (Excel)
keywords: vbaxl10.chm137273
f1_keywords:
- vbaxl10.chm137273
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSub
ms.assetid: bf3d6ea1-46e2-b6d3-66e0-40576db5be2f
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ImSub method (Excel)

Returns the difference of two complex numbers in x + yi or x + yj text format.


## Syntax

_expression_.**ImSub** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber1 - the complex number from which to subtract inumber2.|
| _Arg2_|Required| **Variant**|Inumber2 - the complex number to subtract from inumber1.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
The difference of two complex numbers is &nbsp; ![Formula](../images/awfimsub_ZA06051171.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]