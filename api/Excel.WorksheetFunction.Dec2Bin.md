---
title: WorksheetFunction.Dec2Bin method (Excel)
keywords: vbaxl10.chm137264
f1_keywords:
- vbaxl10.chm137264
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Dec2Bin
ms.assetid: beb1848e-320d-eaef-074e-4df23c46009d
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Dec2Bin method (Excel)

Converts a decimal number to binary.


## Syntax

_expression_.**Dec2Bin** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the decimal integer that you want to convert. If number is negative, valid place values are ignored and **Dec2Bin** returns a 10-character (10-bit) binary number in which the most significant bit is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, **Dec2Bin** uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

## Return value

**String**


## Remarks

If number < -512 or if number > 511, **Dec2Bin** returns the #NUM! error value.
    
If number is nonnumeric, **Dec2Bin** returns the #VALUE! error value.
    
If **Dec2Bin** requires more than places characters, it returns the #NUM! error value.
    
If places is not an integer, it is truncated.
    
If places is nonnumeric, **Dec2Bin** returns the #VALUE! error value.
    
If places is zero or negative, **Dec2Bin** returns the #NUM! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]