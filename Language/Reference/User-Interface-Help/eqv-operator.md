---
title: Eqv operator
keywords: vblr6.chm1008909
f1_keywords:
- vblr6.chm1008909
ms.prod: office
ms.assetid: 6662347b-5229-3bb7-a8f2-d1216094c870
ms.date: 11/19/2018
localization_priority: Normal
---


# Eqv operator

Used to perform a logical equivalence on two [expressions](../../Glossary/vbe-glossary.md#expression).

## Syntax

_result_=_expression1_ **Eqv** _expression2_

The **Eqv** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _expression1_|Required; any expression.|
| _expression2_|Required; any expression.|

## Remarks

If either expression is [Null](../../Glossary/vbe-glossary.md#null), _result_ is also **Null**. When neither expression is **Null**, _result_ is determined according to the following table.

|If  _expression1_ is|And  _expression2_ is|The  _result_ is|
|:-----|:-----|:-----|
|**True**|**True**|**True**|
|**True**|**False**|**False**|
|**False**|**True**|**False**|
|**False**|**False**|**True**|

<br/>

The **Eqv** operator performs a [bitwise comparison](../../Glossary/vbe-glossary.md#bitwise-comparison) of identically positioned bits in two [numeric expressions](../../Glossary/vbe-glossary.md#numeric-expression) and sets the corresponding bit in _result_ according to the following table.

|If bit in _expression1_ is|And bit in _expression2_ is|The _result_ is|
|:-----|:-----|:-----|
|0|0|1|
|0|1|0|
|1|0|0|
|1|1|1|

## Example

This example uses the **Eqv** operator to perform logical equivalence on two expressions.

```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = A > B Eqv B > C    ' Returns True.
MyCheck = B > A Eqv B > C    ' Returns False.
MyCheck = A > B Eqv B > D    ' Returns Null.
MyCheck = A Eqv B    ' Returns -3 (bitwise comparison).
```

## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]