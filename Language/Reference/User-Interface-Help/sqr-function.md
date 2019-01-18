---
title: Sqr function (Visual Basic for Applications)
keywords: vblr6.chm1009029
f1_keywords:
- vblr6.chm1009029
ms.prod: office
ms.assetid: ce2add56-f943-9470-0caa-befda14d124a
ms.date: 12/13/2018
localization_priority: Normal
---


# Sqr function

Returns a **Double** specifying the square root of a number.

## Syntax

**Sqr**(_number_)

The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Double](../../Glossary/vbe-glossary.md#double-data-type) or any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) greater than or equal to zero.

## Example

This example uses the **Sqr** function to calculate the square root of a number.

```vb
Dim MySqr
MySqr = Sqr(4)    ' Returns 2.
MySqr = Sqr(23)    ' Returns 4.79583152331272.
MySqr = Sqr(0)    ' Returns 0.
MySqr = Sqr(-4)    ' Generates a run-time error.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]