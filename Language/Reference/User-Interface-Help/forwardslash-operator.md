---
title: / operator
keywords: vblr6.chm1008852
f1_keywords:
- vblr6.chm1008852
ms.prod: office
ms.assetid: 9d0a9ced-43a0-1c60-29c0-b9ff1062a3e9
ms.date: 11/19/2018
localization_priority: Normal
---

# / operator

Used to divide two numbers and return a floating-point result.

## Syntax

_result_ = _number1_ / _number2_

The **`/`** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _number1_|Required; any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
| _number2_|Required; any numeric expression.|


## Remarks

The data type of _result_ is usually a **Double** or a **Double** variant. The following are exceptions to this rule:

|If|Then _result_ is|
|:-----|:-----|
| Both expressions are **Byte**, **Integer**, or **Single** expressions | A **Single** unless it overflows its legal range, in which case, an error occurs.|
| Both expressions are **Byte**, **Integer**, or **Single** variants | A **Single** variant unless it overflows its legal range, in which case, _result_ is a **Variant** containing a **Double**.| 
| Division involves a **Decimal** and any other data type | A **Decimal** data type.| 

If one or both expressions are **Null** expressions, _result_ is **Null**. Any expression that is **Empty** is treated as 0.

## Example

This example uses the **`/`** operator to perform floating-point division.

```vb
Dim MyValue
MyValue = 10 / 4   ' Returns 2.5.
MyValue = 10 / 3   ' Returns 3.333333.

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]