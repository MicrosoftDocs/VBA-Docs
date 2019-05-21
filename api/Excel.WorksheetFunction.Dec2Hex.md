---
title: WorksheetFunction.Dec2Hex method (Excel)
keywords: vbaxl10.chm137265
f1_keywords:
- vbaxl10.chm137265
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Dec2Hex
ms.assetid: 32e8f754-9d67-1b99-08d3-1eee27237369
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Dec2Hex method (Excel)

Converts a decimal number to hexadecimal.


## Syntax

_expression_.**Dec2Hex** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the decimal integer that you want to convert. If number is negative, places is ignored and **Dec2Hex** returns a 10-character (40-bit) hexadecimal number in which the most significant bit is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, **Dec2Hex** uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

## Return value

**String**


## Remarks

If number < -549,755,813,888 or if number > 549,755,813,887, **Dec2Hex** returns the #NUM! error value.
    
If number is nonnumeric, **Dec2Hex** returns the #VALUE! error value.
    
If **Dec2Hex** requires more than places characters, it returns the #NUM! error value.
    
If places is not an integer, it is truncated.
    
If places is nonnumeric, **Dec2Hex** returns the #VALUE! error value.
    
If places is negative, **Dec2Hex** returns the #NUM! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]