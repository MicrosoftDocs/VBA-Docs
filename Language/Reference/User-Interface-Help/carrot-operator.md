---
title: ^ Operator
keywords: vblr6.chm1008849
f1_keywords:
- vblr6.chm1008849
ms.prod: office
ms.assetid: 9a2f874a-bf55-ae06-cf93-951d774eff0a
ms.date: 06/08/2017
---


# ^ Operator



Used to raise a number to the power of an exponent.

## Syntax

_result_**=**_number_**^**_exponent_
The  **^** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _number_|Required; any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
| _exponent_|Required; any numeric expression.|

## Remarks

A  _number_ can be negative only if _exponent_ is an integer value. When more than one exponentiation is performed in a single[expression](../../Glossary/vbe-glossary.md#expression), the  **^** operator is evaluated as it is encountered from left to right.
Usually, the [data type](../../Glossary/vbe-glossary.md#data-type) of _result_ is a[Double](../../Glossary/vbe-glossary.md#Double) or a[Variant](../../Glossary/vbe-glossary.md#Variant) containing a **Double**. However, if either _number_ or _exponent_ is a[Null](../../Glossary/vbe-glossary.md#Null) expression, _result_ is **Null**.

## Example

This example uses the  **^** operator to raise a number to the power of an exponent.


```vb
Dim MyValue
MyValue = 2 ^ 2    ' Returns 4.
MyValue = 3 ^ 3 ^ 3    ' Returns 19683.
MyValue = (-5) ^ 3    ' Returns -125.


```


