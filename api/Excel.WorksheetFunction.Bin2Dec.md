---
title: WorksheetFunction.Bin2Dec method (Excel)
keywords: vbaxl10.chm137270
f1_keywords:
- vbaxl10.chm137270
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Bin2Dec
ms.assetid: 05a212f7-8330-002f-8bbc-f54550d1276e
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Bin2Dec method (Excel)

Converts a binary number to decimal.


## Syntax

_expression_.**Bin2Dec** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The binary number that you want to convert. Number cannot contain more than 10 characters (10 bits). The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented by using two's-complement notation.|

## Return value

**String**


## Remarks

If number is not a valid binary number, or if number contains more than 10 characters (10 bits), **Bin2Dec** generates an error value.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]