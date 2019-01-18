---
title: Ampersand (&) operator
keywords: vblr6.chm1008852
f1_keywords:
- vblr6.chm1008852
ms.prod: office
ms.assetid: 9abc8a5d-0f53-bbbf-a4b5-7f034609923d
ms.date: 11/19/2018
localization_priority: Normal
---

# & operator

Used to force string concatenation of two [expressions](../../Glossary/vbe-glossary.md#expression).

## Syntax

_result_ = _expression1_ **&** _expression2_

The **&** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _result_|Required; any **String** or **Variant** [variable](../../Glossary/vbe-glossary.md).|
| _expression1_|Required; any expression.|
| _expression2_|Required; any expression.|


## Remarks

If an _expression_ is not a string, it is converted to a **String** variant. The data type of _result_ is **String** if both _expressions_ are string expressions; otherwise, _result_ is a **String** variant. 

If both expressions are **[Null](../../Glossary/vbe-glossary.md#null)**, _result_ is **Null**. However, if only one _expression_ is **Null**, that expression is treated as a zero-length string ("") when concatenated with the other expression. Any expression that is **Empty** is also treated as a zero-length string.

## Example

This example uses the **&** operator to force string concatenation.

```vb
Dim MyStr
MyStr = "Hello" & " World"   ' Returns "Hello World".
MyStr = "Check " & 123 & " Check"   ' Returns "Check 123 Check".
```


## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]