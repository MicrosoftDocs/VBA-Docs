---
title: WorksheetFunction.FVSchedule method (Excel)
keywords: vbaxl10.chm137352
f1_keywords:
- vbaxl10.chm137352
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FVSchedule
ms.assetid: 5a64322c-24b0-baa2-a355-c414fcbe161c
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.FVSchedule method (Excel)

Returns the future value of an initial principal after applying a series of compound interest rates. Use **FVSchedule** to calculate the future value of an investment with a variable or adjustable rate.


## Syntax

_expression_.**FVSchedule** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Principal - the present value.|
| _Arg2_|Required| **Variant**|Schedule - an array of interest rates to apply.|

## Return value

**Double**


## Remarks

The values in schedule can be numbers or blank cells; any other value produces the #VALUE! error value for **FVSchedule**. Blank cells are taken as zeros (no interest). 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]