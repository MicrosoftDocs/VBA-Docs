---
title: WorksheetFunction.Hex2Oct method (Excel)
keywords: vbaxl10.chm137263
f1_keywords:
- vbaxl10.chm137263
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Hex2Oct
ms.assetid: fd1bdc2b-a5bc-e37a-85c4-2275536e2efc
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.Hex2Oct method (Excel)

Converts a hexadecimal number to octal.


## Syntax

_expression_.**Hex2Oct** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the hexadecimal number that you want to convert. Number cannot contain more than 10 characters. The most significant bit of number is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, **Hex2Oct** uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

## Return value

**String**


## Remarks

If number is negative, **Hex2Oct** ignores places and returns a 10-character octal number.
    
If number is negative, it cannot be less than FFE0000000, and if number is positive, it cannot be greater than 1FFFFFFF.
    
If number is not a valid hexadecimal number, **Hex2Oct** returns the #NUM! error value.
    
If **Hex2Oct** requires more than places characters, it returns the #NUM! error value.
    
If places is not an integer, it is truncated.
    
If places is nonnumeric, **Hex2Oct** returns the #VALUE! error value.
    
If places is negative, **Hex2Oct** returns the #NUM! error value.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]