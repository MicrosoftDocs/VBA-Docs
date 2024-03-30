---
title: WorksheetFunction.ImPower method (Excel)
keywords: vbaxl10.chm137275
f1_keywords:
- vbaxl10.chm137275
api_name:
- Excel.WorksheetFunction.ImPower
ms.assetid: 00dfdca2-8609-6719-f666-c8a78998d07e
ms.date: 05/23/2019
ms.localizationpriority: medium
---


# WorksheetFunction.ImPower method (Excel)

Returns a complex number in x + yi or x + yj text format raised to a power.


## Syntax

_expression_.**ImPower** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number that you want to raise to a power.|
| _Arg2_|Required| **Variant**|Number - the power to which you want to raise the complex number.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
If number is nonnumeric, **ImPower** returns the #VALUE! error value.
    
Number can be an integer, fractional, or negative.
    
A complex number raised to a power is calculated as follows:

> ![Screenshot of the complex number formula.](../images/awfimpw1_ZA06051164.gif) 

where: 

> ![Second screenshot of the complex number formula.](../images/awfimpw2_ZA06051165.gif)

and: 

> ![Third screenshot of the complex number formula.](../images/awfimpw3_ZA06051166.gif)

and: 

> ![Fourth screenshot of the complex number formula.](../images/awfimar3_ZA06051155.gif)


    
[!include[Support and feedback](~/includes/feedback-boilerplate.md)]