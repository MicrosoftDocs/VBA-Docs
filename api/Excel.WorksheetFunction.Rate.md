---
title: WorksheetFunction.Rate method (Excel)
keywords: vbaxl10.chm137111
f1_keywords:
- vbaxl10.chm137111
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rate
ms.assetid: 5b412b46-d54a-a36a-a309-c819f2671185
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Rate method (Excel)

Returns the interest rate per period of an annuity. **Rate** is calculated by iteration and can have zero or more solutions. If the successive results of **Rate** do not converge to within 0.0000001 after 20 iterations, **Rate** returns the #NUM! error value.


## Syntax

_expression_.**Rate** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Nper - the total number of payment periods in an annuity.|
| _Arg2_|Required| **Double**|Pmt - the payment made each period and cannot change over the life of the annuity. Typically, pmt includes principal and interest but no other fees or taxes. If pmt is omitted, you must include the fv argument.|
| _Arg3_|Required| **Double**|Pv - the present value&mdash;the total amount that a series of future payments is worth now.|
| _Arg4_|Optional| **Variant**|Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0).|
| _Arg5_|Optional| **Variant**|Type - the number 0 or 1; indicates when payments are due.|
| _Arg6_|Optional| **Variant**|Guess - your guess for what the rate will be.|

## Return value

**Double**


## Remarks

For a complete description of the arguments nper, pmt, pv, fv, and type, see **[Pv](excel.worksheetfunction.pv.md)**.

The following table describes the values that can be used for _Arg5_.

|Set type equal to|If payments are due|
|:-----|:-----|
|0 or omitted|At the end of the period|
|1|At the beginning of the period|

If you omit guess, it is assumed to be 10 percent.
    
If **Rate** does not converge, try different values for guess. **Rate** usually converges if guess is between 0 and 1.
    
Make sure that you are consistent about the units that you use for specifying guess and nper. If you make monthly payments on a four-year loan at 12 percent annual interest, use 12%/12 for guess and 4*12 for nper. If you make annual payments on the same loan, use 12% for guess and 4 for nper.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]