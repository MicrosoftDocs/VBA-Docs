---
title: NPV function (Visual Basic for Applications)
keywords: vblr6.chm1009284
f1_keywords:
- vblr6.chm1009284
ms.prod: office
ms.assetid: 9f444237-9f5a-834d-1aec-a2d016dfb325
ms.date: 12/13/2018
localization_priority: Normal
---


# NPV function

Returns a [Double](../../Glossary/vbe-glossary.md#double-data-type) specifying the net present value of an investment based on a series of periodic cash flows (payments and receipts) and a discount rate.

## Syntax

**NPV**(_rate_, _values_( ))

The **NPV** function has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_rate_|Required. **Double** specifying discount rate over the length of the period, expressed as a decimal.|
|_values_(&nbsp;) |Required. [Array](../../Glossary/vbe-glossary.md#array) of **Double** specifying cash flow values. The array must contain at least one negative value (a payment) and one positive value (a receipt).|

## Remarks

The net present value of an investment is the current value of a future series of payments and receipts.

The **NPV** function uses the order of values within the array to interpret the order of payments and receipts. Be sure to enter your payment and receipt values in the correct sequence.

The **NPV** investment begins one period before the date of the first cash flow value and ends with the last cash flow value in the array.

The net present value calculation is based on future cash flows. If your first cash flow occurs at the beginning of the first period, the first value must be added to the value returned by **NPV** and must not be included in the cash flow values of _values_( ).

The **NPV** function is similar to the **[PV](pv-function.md)** function (present value) except that the **PV** function allows cash flows to begin either at the end or the beginning of a period. Unlike the variable **NPV** cash flow values, **PV** cash flows must be fixed throughout the investment.

## Example

This example uses the **NPV** function to return the net present value for a series of cash flows contained in the array `Values()`. `RetRate` represents the fixed internal rate of return.

```vb
Dim Fmt, Guess, RetRate, NetPVal, Msg
Static Values(5) As Double    ' Set up array.
Fmt = "###,##0.00"    ' Define money format.
Guess = .1    ' Guess starts at 10 percent.
RetRate = .0625    ' Set fixed internal rate.
Values(0) = -70000    ' Business start-up costs.
' Positive cash flows reflecting income for four successive years.
Values(1) = 22000 : Values(2) = 25000
Values(3) = 28000 : Values(4) = 31000
NetPVal = NPV(RetRate, Values())    ' Calculate net present value.
Msg = "The net present value of these cash flows is "
Msg = Msg & Format(NetPVal, Fmt) & "."
MsgBox Msg    ' Display net present value.
```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]