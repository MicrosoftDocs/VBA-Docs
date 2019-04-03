---
title: WorksheetFunction.ImLn method (Excel)
keywords: vbaxl10.chm137278
f1_keywords:
- vbaxl10.chm137278
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImLn
ms.assetid: a2542e7d-f46b-bb01-67a6-655a92f782c9
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.ImLn method (Excel)

Returns the natural logarithm of a complex number in x + yi or x + yj text format.


## Syntax

_expression_. `ImLn`( `_Arg1_` )

_expression_ A variable that represents a [WorksheetFunction](Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the natural logarithm.|

## Return value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The natural logarithm of a complex number is:
![Formula](../images/awfimln_ZA06051162.gif)where: 
![Formula](../images/awfimar3_ZA06051155.gif)


    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]