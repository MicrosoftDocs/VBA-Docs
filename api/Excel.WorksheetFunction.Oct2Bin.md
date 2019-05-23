---
title: WorksheetFunction.Oct2Bin method (Excel)
keywords: vbaxl10.chm137267
f1_keywords:
- vbaxl10.chm137267
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Oct2Bin
ms.assetid: a11c26e2-1320-f76f-547e-fa9e0ac20087
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Oct2Bin method (Excel)

Converts an octal number to binary.


## Syntax

_expression_.**Oct2Bin** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the octal number that you want to convert. Number may not contain more than 10 characters. The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, **Oct2Bin** uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

## Return value

**String**


## Remarks

If number is negative, **Oct2Bin** ignores places and returns a 10-character binary number.
    
If number is negative, it cannot be less than 7777777000, and if number is positive, it cannot be greater than 777.
    
If number is not a valid octal number, **Oct2Bin** returns the #NUM! error value.
    
If **Oct2Bin** requires more than places characters, it returns the #NUM! error value.
    
If places is not an integer, it is truncated.
    
If places is nonnumeric, **Oct2Bin** returns the #VALUE! error value.
    
If places is negative, **Oct2Bin** returns the #NUM! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]