---
title: WorksheetFunction.ImAbs method (Excel)
keywords: vbaxl10.chm137276
f1_keywords:
- vbaxl10.chm137276
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImAbs
ms.assetid: 630fc586-8899-59e2-dde9-629c08f2b8eb
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.ImAbs method (Excel)

Returns the absolute value (modulus) of a complex number in x + yi or x + yj text format.


## Syntax

_expression_. `ImAbs`( `_Arg1_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the absolute value.|

## Return value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The absolute value of a complex number is:
![Formula](../images/awfimabs_ZA06051152.gif)where: z = x + yi 
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]