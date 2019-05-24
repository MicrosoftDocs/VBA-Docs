---
title: WorksheetFunction.Ppmt method (Excel)
keywords: vbaxl10.chm137141
f1_keywords:
- vbaxl10.chm137141
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Ppmt
ms.assetid: 9b088c90-9db6-b043-ee3e-0f9709a06475
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Ppmt method (Excel)

Returns the payment on the principal for a given period for an investment based on periodic, constant payments and a constant interest rate.


## Syntax

_expression_.**Ppmt** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Rate - the interest rate per period.|
| _Arg2_|Required| **Double**|Per - the period and must be in the range 1 to nper.|
| _Arg3_|Required| **Double**|Nper - the total number of payment periods in an annuity.|
| _Arg4_|Required| **Double**|Pv - the present value&mdash;the total amount that a series of future payments is worth now.|
| _Arg5_|Optional| **Variant**|Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (zero), that is, the future value of a loan is 0.|
| _Arg6_|Optional| **Variant**|Type - the number 0 or 1 and indicates when payments are due.|

## Return value

**Double**


## Remarks

For a more complete description of the arguments in **Ppmt**, see the **[Pv](excel.worksheetfunction.pv.md)** function.

The following table describes the values that can be used for _Arg6_.

|Set type equal to|If payments are due|
|:-----|:-----|
|0 or omitted|At the end of the period|
|1|At the beginning of the period|

Make sure that you are consistent about the units that you use for specifying rate and nper. If you make monthly payments on a four-year loan at 12 percent annual interest, use 12%/12 for rate and 4*12 for nper. If you make annual payments on the same loan, use 12% for rate and 4 for nper.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]