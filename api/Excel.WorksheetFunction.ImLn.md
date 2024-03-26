---
title: WorksheetFunction.ImLn method (Excel)
keywords: vbaxl10.chm137278
f1_keywords:
- vbaxl10.chm137278
api_name:
- Excel.WorksheetFunction.ImLn
ms.assetid: a2542e7d-f46b-bb01-67a6-655a92f782c9
ms.date: 05/23/2019
ms.localizationpriority: medium
---


# WorksheetFunction.ImLn method (Excel)

Returns the natural logarithm of a complex number in x + yi or x + yj text format.


## Syntax

_expression_.**ImLn** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the natural logarithm.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
The natural logarithm of a complex number is &nbsp; ![Screenshot of the natural logarithm of a complex number formula.](../images/awfimln_ZA06051162.gif) &nbsp; where &nbsp; ![Screenshot of the theta value formula.](../images/awfimar3_ZA06051155.gif)


    

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]