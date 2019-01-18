---
title: WorksheetFunction.ImSin method (Excel)
keywords: vbaxl10.chm137281
f1_keywords:
- vbaxl10.chm137281
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSin
ms.assetid: 1c9d4442-954e-3943-747f-647f49b4257c
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.ImSin method (Excel)

Returns the sine of a complex number in x + yi or x + yj text format.


## Syntax

_expression_. `ImSin`( `_Arg1_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the sine.|

## Return value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The sine of a complex number is:
![Formula](../images/awfimsin_ZA06051167.gif)


    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]