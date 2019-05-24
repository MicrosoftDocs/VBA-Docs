---
title: WorksheetFunction.Oct2Dec method (Excel)
keywords: vbaxl10.chm137269
f1_keywords:
- vbaxl10.chm137269
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Oct2Dec
ms.assetid: 08694db9-136b-9bfe-7939-436f4773bffb
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Oct2Dec method (Excel)

Converts an octal number to decimal.


## Syntax

_expression_.**Oct2Dec** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the octal number that you want to convert. Number may not contain more than 10 octal characters (30 bits). The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|

## Return value

**String**


## Remarks

If number is not a valid octal number, **Oct2Dec** returns the #NUM! error value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]