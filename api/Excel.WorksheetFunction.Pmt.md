---
title: WorksheetFunction.Pmt method (Excel)
keywords: vbaxl10.chm137110
f1_keywords:
- vbaxl10.chm137110
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Pmt
ms.assetid: ef383e8e-7fca-2818-cdaa-d758f2e8536d
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Pmt method (Excel)

Calculates the payment for a loan based on constant payments and a constant interest rate.


## Syntax

_expression_.**Pmt** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Rate - the interest rate for the loan.|
| _Arg2_|Required| **Double**|Nper - the total number of payments for the loan.|
| _Arg3_|Required| **Double**|Pv - the present value, or the total amount that a series of future payments is worth now; also known as the principal.|
| _Arg4_|Optional| **Variant**|Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (zero), that is, the future value of a loan is 0.|
| _Arg5_|Optional| **Variant**|Type - the number 0 (zero) or 1; indicates when payments are due.|

## Return value

**Double**


## Remarks

For a more complete description of the arguments in **Pmt**, see the **[Pv](excel.worksheetfunction.pv.md)** function.

The following table describes the values that can be used for _Arg5_.

|Set type equal to|If payments are due|
|:-----|:-----|
|0 or omitted|At the end of the period|
|1|At the beginning of the period|

The payment returned by **Pmt** includes principal and interest but no taxes, reserve payments, or fees sometimes associated with loans.
    
Make sure that you are consistent about the units that you use for specifying rate and nper. If you make monthly payments on a four-year loan at an annual interest rate of 12 percent, use 12%/12 for rate and 4*12 for nper. If you make annual payments on the same loan, use 12 percent for rate and 4 for nper.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]