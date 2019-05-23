---
title: WorksheetFunction.Fv method (Excel)
keywords: vbaxl10.chm137108
f1_keywords:
- vbaxl10.chm137108
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Fv
ms.assetid: 0f2cedc5-2f10-0ad1-b140-cdbbfa6af8ce
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Fv method (Excel)

Returns the future value of an investment based on periodic, constant payments and a constant interest rate.


## Syntax

_expression_.**Fv** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Rate - the interest rate per period.|
| _Arg2_|Required| **Double**|Nper - the total number of payment periods in an annuity.|
| _Arg3_|Required| **Double**|Pmt - the payment made each period; it cannot change over the life of the annuity. Typically, pmt contains principal and interest but no other fees or taxes. If pmt is omitted, you must include the pv argument.|
| _Arg4_|Optional| **Variant**|Pv - the present value, or the lump-sum amount that a series of future payments is worth right now. If pv is omitted, it is assumed to be 0 (zero), and you must include the pmt argument.|
| _Arg5_|Optional| **Variant**|Type - the number 0 or 1 and indicates when payments are due. If type is omitted, it is assumed to be 0.|

## Return value

**Double**


## Remarks

For a more complete description of the arguments in **Fv** and for more information about annuity functions, see **[Pv](excel.worksheetfunction.pv.md)**.

The following table describes the values that can be used for _Arg5_.

|Set type equal to|If payments are due|
|:-----|:-----|
|0|At the end of the period|
|1|At the beginning of the period|

Make sure that you are consistent about the units you use for specifying rate and nper. If you make monthly payments on a four-year loan at 12 percent annual interest, use 12%/12 for rate and 4*12 for nper. If you make annual payments on the same loan, use 12% for rate and 4 for nper.
    
For all the arguments, cash you pay out, such as deposits to savings, is represented by negative numbers; cash you receive, such as dividend checks, is represented by positive numbers.
    





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]