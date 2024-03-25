---
title: WorksheetFunction.ImArgument method (Excel)
keywords: vbaxl10.chm137284
f1_keywords:
- vbaxl10.chm137284
api_name:
- Excel.WorksheetFunction.ImArgument
ms.assetid: ac1e721a-edfe-0287-afa1-509f5c437cd8
ms.date: 05/23/2019
ms.localizationpriority: medium
---


# WorksheetFunction.ImArgument method (Excel)

Returns the argument ![Screenshot of the theta symbol.](../images/theta_ZA06052070.gif) (theta), an angle expressed in radians, such that:

> ![Screenshot of the angle formula.](../images/awfimar1_ZA06051153.gif)


## Syntax

_expression_.**ImArgument** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber is a complex number for which you want the argument theta.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
**ImArgument** is calculated as follows:

> ![Screenshot of Im Argument formula.](../images/awfimar2_ZA06051154.gif) &nbsp; where &nbsp; ![Screenshot of the theta value formula.](../images/awfimar3_ZA06051155.gif) &nbsp; and z = x + yi
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]