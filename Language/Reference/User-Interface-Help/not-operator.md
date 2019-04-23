---
title: Not operator
keywords: vblr6.chm1008981
f1_keywords:
- vblr6.chm1008981
ms.prod: office
ms.assetid: e5ae5a73-4f34-0071-ee67-98e4ca519748
ms.date: 01/02/2019
localization_priority: Normal
---


# Not operator

Used to perform logical negation on an [expression](../../Glossary/vbe-glossary.md#expression).

## Syntax

_result_ = **Not** _expression_

The **Not** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _expression_|Required; any expression.|

## Remarks

The following table illustrates how _result_ is determined.

|If _expression_ is|Then _result_ is|
|:-----|:-----|
|**True**|**False**|
|**False**|**True**|
|[Null](../../Glossary/vbe-glossary.md#null)|**Null**|

In addition, the **Not** operator inverts the bit values of any variable and sets the corresponding bit in _result_ according to the following table.

|If bit in _expression_ is|Then bit in _result_ is|
|:-----|:-----|
|0|1|
|1|0|

## Example

This example uses the **Not** operator to perform logical negation on an expression.

```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = Not(A > B)    ' Returns False.
MyCheck = Not(B > A)    ' Returns True.
MyCheck = Not(C > D)    ' Returns Null.
MyCheck = Not A    ' Returns -11 (bitwise comparison).

```

## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]