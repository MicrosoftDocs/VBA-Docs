---
title: Mod operator
keywords: vblr6.chm1008976
f1_keywords:
- vblr6.chm1008976
ms.prod: office
ms.assetid: cc1afd5d-ea12-a1df-3ffe-0d58f4d1e0ac
ms.date: 11/19/2018
localization_priority: Normal
---


# Mod operator

Used to divide two numbers and return only the remainder.

## Syntax

_result_ = _number1_ **Mod** _number2_

The **Mod** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _number1_|Required; any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
| _number2_|Required; any numeric expression.|

## Remarks

The modulus, or remainder, operator divides _number1_ by _number2_ (rounding floating-point numbers to integers) and returns only the remainder as _result_. For example, in the following [expression](../../Glossary/vbe-glossary.md#expression), A (_result_) equals 5.

```vb
A = 19 Mod 6.7
```

Usually, the [data type](../../Glossary/vbe-glossary.md#data-type) of _result_ is a [Byte](../../Glossary/vbe-glossary.md#byte-data-type), **Byte** variant, [Integer](../../Glossary/vbe-glossary.md#integer-data-type), **Integer** variant, [Long](../../Glossary/vbe-glossary.md#long-data-type), or [Variant](../../Glossary/vbe-glossary.md#variant-data-type) containing a **Long**, regardless of whether or not _result_ is a whole number. Any fractional portion is truncated. 

However, if any expression is [Null](../../Glossary/vbe-glossary.md#null), _result_ is **Null**. Any expression that is [Empty](../../Glossary/vbe-glossary.md#empty) is treated as 0.

## Example

This example uses the **Mod** operator to divide two numbers and return only the remainder. If either number is a floating-point number, it is first rounded to an integer.

```vb
Dim MyResult
MyResult = 10 Mod 5    ' Returns 0.
MyResult = 10 Mod 3    ' Returns 1.
MyResult = 12 Mod 4.3    ' Returns 0.
MyResult = 12.6 Mod 5    ' Returns 3.
```

## See also

- [Mod operator examples (previous versions)](https://docs.microsoft.com/previous-versions/office/office-10/aa263659(v=office.10))
- [Data types](data-type-summary.md)
- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
