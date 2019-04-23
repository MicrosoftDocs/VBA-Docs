---
title: Or operator
keywords: vblr6.chm1008993
f1_keywords:
- vblr6.chm1008993
ms.prod: office
ms.assetid: 3b0e4886-2f84-1296-9428-69338d033c6c
ms.date: 11/19/2018
localization_priority: Normal
---


# Or operator

Used to perform a logical disjunction on two [expressions](../../Glossary/vbe-glossary.md#expression).

## Syntax

_result_ = _expression1_ **Or** _expression2_

The **Or** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any numeric [variable](../../Glossary/vbe-glossary.md#variable).|
| _expression1_|Required; any expression.|
| _expression2_|Required; any expression.|

## Remarks

If either or both expressions evaluate to **True**, _result_ is **True**. The following table illustrates how _result_ is determined.

|If _expression1_ is|And _expression2_ is|Then _result_ is|
|:-----|:-----|:-----|
|**True**|**True**|**True**|
|**True**|**False**|**True**|
|**True**|[Null](../../Glossary/vbe-glossary.md#null)|**True**|
|**False**|**True**|**True**|
|**False**|**False**|**False**|
|**False**|**Null**|**Null**|
|**Null**|**True**|**True**|
|**Null**|**False**|**Null**|
|**Null**|**Null**|**Null**|

<br/>

The **Or** operator also performs a [bitwise comparison](../../Glossary/vbe-glossary.md#bitwise-comparison) of identically positioned bits in two [numeric expressions](../../Glossary/vbe-glossary.md#numeric-expression) and sets the corresponding bit in _result_ according to the following table.

|If bit in _expression1_ is|And bit in _expression2_ is|Then _result_ is|
|:-----|:-----|:-----|
|0|0|0|
|0|1|1|
|1|0|1|
|1|1|1|

## Example

This example uses the **Or** operator to perform logical disjunction on two expressions.

```vb
Dim A, B, C, D, MyCheck
A = 10: B = 8: C = 6: D = Null    ' Initialize variables.
MyCheck = A > B Or B > C    ' Returns True.
MyCheck = B > A Or B > C    ' Returns True.
MyCheck = A > B Or B > D    ' Returns True.
MyCheck = B > D Or B > A    ' Returns Null.
MyCheck = A Or B    ' Returns 10 (bitwise comparison).

```


## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]