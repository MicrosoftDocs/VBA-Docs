---
title: WorksheetFunction.Quotient method (Excel)
keywords: vbaxl10.chm137294
f1_keywords:
- vbaxl10.chm137294
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quotient
ms.assetid: 33a057f8-dbb7-0f0e-fabd-ebdd4d471159
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Quotient method (Excel)

Returns the integer portion of a division. Use this function when you want to discard the remainder of a division.


## Syntax

_expression_.**Quotient** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Numerator - the dividend.|
| _Arg2_|Required| **Variant**|Denominator - the divisor.|

## Return value

**Double**


## Remarks

If either argument is nonnumeric, **Quotient** returns the #VALUE! error value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]