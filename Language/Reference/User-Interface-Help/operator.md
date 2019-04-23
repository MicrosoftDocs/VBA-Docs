---
title: "* operator"
keywords: vblr6.chm1008844
f1_keywords:
- vblr6.chm1008844
ms.prod: office
ms.assetid: f45e939e-ff1d-c152-ad82-099e8f00ee8c
ms.date: 11/19/2018
localization_priority: Normal
---


# * Operator

Used to multiply two numbers.

## Syntax

_result_ = _number1_ * _number2_

The **\*** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _number1_|Required; any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
| _number2_|Required; any numeric expression.|

## Remarks

The [data type](../../Glossary/vbe-glossary.md#data-type) of _result_ is usually the same as that of the most precise[expression](../../Glossary/vbe-glossary.md#expression). The order of precision, from least to most precise, is [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), and [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type). 

The following are exceptions to this order:

|If|Then _result_ is|
|:-----|:-----|
|Multiplication involves a **Single** and a **Long**|Converted to a **Double**.|
|The data type of _result_ is a **Long**, **Single**, or **Date** variant that overflows its legal range |Converted to a **Variant** containing a **Double**.|
|The data type of _result_ is a **Byte** variant that overflows its legal range |Converted to an **Integer** variant.|
|The data type of _result_ is an **Integer** variant that overflows its legal range |Converted to a **Long** variant.|

If one or both expressions are [Null](../../Glossary/vbe-glossary.md#null) expressions, _result_ is **Null**. If an expression is [Empty](../../Glossary/vbe-glossary.md#empty), it is treated as 0.

> [!NOTE] 
> The order of precision used by multiplication is not the same as the order of precision used by addition and subtraction.


## Example

This example uses the **\*** operator to multiply two numbers.

```vb
Dim MyValue
MyValue = 2 * 2    ' Returns 4.
MyValue = 459.35 * 334.90     ' Returns 153836.315.

```

## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]