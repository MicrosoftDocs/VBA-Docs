---
title: WorksheetFunction.ImLog2 method (Excel)
keywords: vbaxl10.chm137279
f1_keywords:
- vbaxl10.chm137279
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImLog2
ms.assetid: 7eb55cd5-fec2-c110-981b-81c55b241900
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ImLog2 method (Excel)

Returns the base-2 logarithm of a complex number in x + yi or x + yj text format.


## Syntax

_expression_.**ImLog2** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the base-2 logarithm.|

## Return value

**String**


## Remarks

Use the **[Complex](excel.worksheetfunction.complex.md)** method to convert real and imaginary coefficients into a complex number.
    
The base-2 logarithm of a complex number can be calculated from the natural logarithm as follows:

> ![Formula](../images/awfimlg2_ZA06051161.gif)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]