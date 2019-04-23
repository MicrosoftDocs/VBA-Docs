---
title: \ operator
keywords: vblr6.chm1008852
f1_keywords:
- vblr6.chm1008852
ms.prod: office
ms.assetid: ec04fbea-3cc1-4b9b-b1e0-008980ba404e
ms.date: 11/19/2018
localization_priority: Normal
---

# \ operator

Used to divide two numbers and return an integer result.

## Syntax

_result_ = _number1_ \ _number2_

The **`\`** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _number1_|Required; any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
| _number2_|Required; any numeric expression.|


## Remarks

Before division is performed, the numeric expressions are rounded to **Byte**, **Integer**, or **Long** expressions.

Usually, the data type of _result_ is a **Byte**, **Byte** variant, **Integer**, **Integer** variant, **Long**, or **Long** variant, regardless of whether _result_ is a whole number. 

Any fractional portion is truncated. However, if any expression is **Null**, _result_ is **Null**. Any expression that is **Empty** is treated as 0.

## Example

This example uses the **`\`** operator to perform integer division.

```vb
Dim MyValue
MyValue = 11 \ 4   ' Returns 2.
MyValue = 9 \ 3   ' Returns 3. 
MyValue = 100 \ 3   ' Returns 33.

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]