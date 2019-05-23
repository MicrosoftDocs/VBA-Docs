---
title: WorksheetFunction.Irr method (Excel)
keywords: vbaxl10.chm137113
f1_keywords:
- vbaxl10.chm137113
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Irr
ms.assetid: 306de022-0082-9757-9b63-262c7e2e55f4
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.Irr method (Excel)

Returns the internal rate of return for a series of cash flows represented by the numbers in values. These cash flows do not have to be even, as they would be for an annuity. However, the cash flows must occur at regular intervals, such as monthly or annually. The internal rate of return is the interest rate received for an investment consisting of payments (negative values) and income (positive values) that occur at regular periods.


## Syntax

_expression_.**Irr** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Values - an array or a reference to cells that contain numbers for which you want to calculate the internal rate of return.|
| _Arg2_|Optional| **Variant**|Guess - a number that you guess is close to the result of **Irr**.|

## Return value

**Double**


## Remarks

Values must contain at least one positive value and one negative value to calculate the internal rate of return.
    
**Irr** uses the order of values to interpret the order of cash flows. Be sure to enter your payment and income values in the sequence that you want.
    
If an array or reference argument contains text, logical values, or empty cells, those values are ignored.
    
Microsoft Excel uses an iterative technique for calculating **Irr**. Starting with guess, **Irr** cycles through the calculation until the result is accurate within 0.00001 percent. If **Irr** can't find a result that works after 20 tries, the #NUM! error value is returned.
    
In most cases, you do not need to provide guess for the **Irr** calculation. If guess is omitted, it is assumed to be 0.1 (10 percent).
    
If **Irr** gives the #NUM! error value, or if the result is not close to what you expected, try again with a different value for guess.
    
**Irr** is closely related to **[Npv](excel.worksheetfunction.npv.md)**, the net present value function. The rate of return calculated by **Irr** is the interest rate corresponding to a 0 (zero) net present value. The following formula demonstrates how **Npv** and **Irr** are related:

`NPV(IRR(B1:B6),B1:B6)` equals 3.60E-08 [Within the accuracy of the **Irr** calculation, the value 3.60E-08 is effectively 0 (zero).]




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]