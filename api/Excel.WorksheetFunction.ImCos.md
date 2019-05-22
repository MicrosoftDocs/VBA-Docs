---
title: WorksheetFunction.ImCos method (Excel)
keywords: vbaxl10.chm137282
f1_keywords:
- vbaxl10.chm137282
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImCos
ms.assetid: 959ac671-64e4-ac72-9421-d7074bd5d4a8
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ImCos method (Excel)

Returns the cosine of a complex number in x + yi or x + yj text format.


## Syntax

_expression_.**ImCos** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the cosine.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
If inumber is a logical value, **ImCos** returns the #VALUE! error value.
    
The cosine of a complex number is &nbsp; ![Formula](../images/awfimcos_ZA06051157.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]