---
title: WorksheetFunction.Dec2Oct method (Excel)
keywords: vbaxl10.chm137266
f1_keywords:
- vbaxl10.chm137266
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Dec2Oct
ms.assetid: 2aac7d4d-57ef-0d8f-1432-62e98ddc1c41
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Dec2Oct method (Excel)

Converts a decimal number to octal.


## Syntax

_expression_.**Dec2Oct** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the decimal integer that you want to convert. If number is negative, places is ignored and **Dec2Oct** returns a 10-character (30-bit) octal number in which the most significant bit is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, **Dec2Oct** uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

## Return value

**String**


## Remarks

If number < -536,870,912 or if number > 536,870,911, **Dec2Oct** returns the #NUM! error value.
    
If number is nonnumeric, **Dec2Oct** returns the #VALUE! error value.
    
If **Dec2Oct** requires more than places characters, it returns the #NUM! error value.
    
If places is not an integer, it is truncated.
    
If places is nonnumeric, **Dec2Oct** returns the #VALUE! error value.
    
If places is negative, **Dec2Oct** returns the #NUM! error value.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]