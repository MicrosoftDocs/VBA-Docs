---
title: MIRR function (Visual Basic for Applications)
keywords: vblr6.chm1009283
f1_keywords:
- vblr6.chm1009283
ms.prod: office
ms.assetid: defc1846-572b-ae88-a845-f732b0a2a15a
ms.date: 12/13/2018
localization_priority: Normal
---


# MIRR function

Returns a [Double](../../Glossary/vbe-glossary.md#double-data-type) specifying the modified internal rate of return for a series of periodic cash flows (payments and receipts).

## Syntax

**MIRR**(_values_( ), _finance\_rate_, _reinvest\_rate_)

<br/>

The **MIRR** function has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_values_( )|Required. [Array](../../Glossary/vbe-glossary.md#array) of **Double** specifying cash flow values. The array must contain at least one negative value (a payment) and one positive value (a receipt).|
|_finance\_rate_|Required. **Double** specifying interest rate paid as the cost of financing.|
|_reinvest\_rate_|Required. **Double** specifying interest rate received on gains from cash reinvestment.|

## Remarks

The modified internal rate of return is the internal rate of return when payments and receipts are financed at different rates. The **MIRR** function takes into account both the cost of the investment (_finance\_rate_) and the interest rate received on reinvestment of cash (_reinvest\_rate_).

The _finance\_rate_ and _reinvest\_rate_ [arguments](../../Glossary/vbe-glossary.md#argument) are percentages expressed as decimal values. For example, 12 percent is expressed as 0.12.

The **MIRR** function uses the order of values within the array to interpret the order of payments and receipts. Be sure to enter your payment and receipt values in the correct sequence.

## Example

This example uses the **MIRR** function to return the modified internal rate of return for a series of cash flows contained in the array `Values()`. `LoanAPR` represents the financing interest, and `InvAPR` represents the interest rate received on reinvestment.

```vb
Dim LoanAPR, InvAPR, Fmt, RetRate, Msg
Static Values(5) As Double    ' Set up array.
LoanAPR = .1    ' Loan rate.
InvAPR = .12    ' Reinvestment rate.
Fmt = "#0.00"    ' Define money format.
Values(0) = -70000    ' Business start-up costs.
' Positive cash flows reflecting income for four successive years.
Values(1) = 22000 : Values(2) = 25000
Values(3) = 28000 : Values(4) = 31000
RetRate = MIRR(Values(), LoanAPR, InvAPR)    ' Calculate internal rate.
Msg = "The modified internal rate of return for these five cash flows is"
Msg = Msg & Format(Abs(RetRate) * 100, Fmt) & "%."
MsgBox Msg    ' Display internal return rate.
        
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]