---
title: Mod Operator
keywords: vblr6.chm1008976
f1_keywords:
- vblr6.chm1008976
ms.prod: office
ms.assetid: cc1afd5d-ea12-a1df-3ffe-0d58f4d1e0ac
ms.date: 06/08/2017
---


# Mod Operator



Used to divide two numbers and return only the remainder.

## Syntax

_result_**=**_number1_**Mod**_number2_
The  **Mod** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _number1_|Required; any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
| _number2_|Required; any numeric expression.|

## Remarks

The modulus, or remainder, operator divides  _number1_ by _number2_ (rounding floating-point numbers to integers) and returns only the remainder as _result_. For example, in the following[expression](../../Glossary/vbe-glossary.md#expression), A ( _result_ ) equals 5.
Usually, the [data type](../../Glossary/vbe-glossary.md#data-type) of _result_ is a[Byte](../../Glossary/vbe-glossary.md#Byte),  **Byte** variant,[Integer](../../Glossary/vbe-glossary.md#Integer),  **Integer** variant,[Long](../../Glossary/vbe-glossary.md#Long), or [Variant](../../Glossary/vbe-glossary.md#Variant) containing a **Long**, regardless of whether or not _result_ is a whole number. Any fractional portion is truncated. However, if any expression is[Null](../../Glossary/vbe-glossary.md#Null),  _result_ is **Null**. Any expression that is[Empty](../../Glossary/vbe-glossary.md#Empty) is treated as 0.

## Example

This example uses the  **Mod** operator to divide two numbers and return only the remainder. If either number is a floating-point number, it is first rounded to an integer.


```vb
Dim MyResult
MyResult = 10 Mod 5    ' Returns 0.
MyResult = 10 Mod 3    ' Returns 1.
MyResult = 12 Mod 4.3    ' Returns 0.
MyResult = 12.6 Mod 5    ' Returns 3.
```


