---
title: WorksheetFunction.Hex2Dec method (Excel)
keywords: vbaxl10.chm137262
f1_keywords:
- vbaxl10.chm137262
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Hex2Dec
ms.assetid: e2e0614c-583e-8a1f-b852-683c119d5a5a
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.Hex2Dec method (Excel)

Converts a hexadecimal number to decimal.


## Syntax

_expression_.**Hex2Dec** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the hexadecimal number that you want to convert. Number cannot contain more than 10 characters (40 bits). The most significant bit of number is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|

## Return value

**String**


## Remarks

If number is not a valid hexadecimal number, **Hex2Dec** returns the #NUM! error value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]