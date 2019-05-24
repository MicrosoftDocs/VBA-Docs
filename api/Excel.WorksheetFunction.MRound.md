---
title: WorksheetFunction.MRound method (Excel)
keywords: vbaxl10.chm137299
f1_keywords:
- vbaxl10.chm137299
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MRound
ms.assetid: 66a8641e-3797-43a4-2b4e-a4c555391c72
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.MRound method (Excel)

Returns a number rounded to the desired multiple.


## Syntax

_expression_.**MRound** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the value to round.|
| _Arg2_|Required| **Variant**|Multiple - the multiple to which you want to round number.|

## Return value

**Double**


## Remarks

**MRound** rounds up, away from zero, if the remainder of dividing number by multiple is greater than or equal to half the value of multiple.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]