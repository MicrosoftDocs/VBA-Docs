---
title: WorksheetFunction.Odd method (Excel)
keywords: vbaxl10.chm137202
f1_keywords:
- vbaxl10.chm137202
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Odd
ms.assetid: 28a30d51-ba7b-f7b4-55a5-39b85f7f4cd7
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Odd method (Excel)

Returns number rounded up to the nearest odd integer.


## Syntax

_expression_.**Odd** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the value to round.|

## Return value

**Double**


## Remarks

If number is nonnumeric, **Odd** returns the #VALUE! error value.
    
Regardless of the sign of number, a value is rounded up when adjusted away from zero. If number is an odd integer, no rounding occurs.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]