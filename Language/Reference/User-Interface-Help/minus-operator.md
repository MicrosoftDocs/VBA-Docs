---
title: Minus (-) operator
keywords: vblr6.chm1008852
f1_keywords:
- vblr6.chm1008852
ms.prod: office
ms.assetid: db8adf52-58a9-9ba8-b5ae-5cdf59f56598
ms.date: 11/19/2018
localization_priority: Normal
---

# - operator

Used to find the difference between two numbers or to indicate the negative value of a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).

## Syntax

### Syntax 1

_result_ = _number1_ **-** _number2_

### Syntax 2

**`-`** number

The **`-`** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_ | Required; any numeric variable.|
| _number_ | Required; any numeric expression.|
| _number1_ | Required; any numeric expression.|
| _number2_ | Required; any numeric expression.|

## Remarks

In Syntax 1, the **`-`** operator is the arithmetic subtraction operator used to find the difference between two numbers. In Syntax 2, the **`-`** operator is used as the unary negation operator to indicate the negative value of an expression.

The data type of _result_ is usually the same as that of the most precise expression. The order of precision, from least to most precise, is [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), and [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type). The following are exceptions to this order:


|If|Then _result_ is|
|:-----|:-----|
| Subtraction involves a **Single** and a **Long** | Converted to a **Double**.|
| The data type of _result_ is a **Long**, **Single**, or **Date** variant that overflows its legal range | Converted to a **Variant** containing a **Double**.|
| The data type of _result_ is a **Byte** variant that overflows its legal range | Converted to an **Integer** variant.| 
| The data type of _result_ is an **Integer** variant that overflows its legal range | Converted to a **Long** variant.| 
| Subtraction involves a **Date** and any other data type | A **Date**.| 
| Subtraction involves two **Date** expressions | A **Double**.| 

If one or both expressions are **Null** expressions, _result_ is **Null**. If an expression is **Empty**, it is treated as 0.

> [!NOTE] 
> The order of precision used by addition and subtraction is not the same as the order of precision used by multiplication.

## Example

This example uses the **-** operator to calculate the difference between two numbers.

```vb 
Dim MyResult
MyResult = 4 - 2   ' Returns 2.
MyResult = 459.35 - 334.90   ' Returns 124.45.

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]