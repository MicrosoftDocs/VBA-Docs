---
title: WorksheetFunction.Oct2Hex method (Excel)
keywords: vbaxl10.chm137268
f1_keywords:
- vbaxl10.chm137268
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Oct2Hex
ms.assetid: eee1bb9b-6b79-aea1-453d-4e2e69b16934
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Oct2Hex method (Excel)

Converts an octal number to hexadecimal.


## Syntax

_expression_.**Oct2Hex** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the octal number that you want to convert. Number may not contain more than 10 octal characters (30 bits). The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, **Oct2Hex** uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

## Return value

**String**


## Remarks

If number is negative, **Oct2Hex** ignores places and returns a 10-character hexadecimal number.
    
If number is not a valid octal number, **Oct2Hex** returns the #NUM! error value.
    
If **Oct2Hex** requires more than places characters, it returns the #NUM! error value.
    
If places is not an integer, it is truncated.
    
If places is nonnumeric, **Oct2Hex** returns the #VALUE! error value.
    
If places is negative, **Oct2Hex** returns the #NUM! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]