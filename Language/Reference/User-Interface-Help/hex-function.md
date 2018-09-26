---
title: Hex Function
keywords: vblr6.chm1011362
f1_keywords:
- vblr6.chm1011362
ms.prod: office
ms.assetid: 79a403a9-61af-0991-8f13-60c1033f158a
ms.date: 06/08/2017
---


# Hex Function



Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) representing the hexadecimal value of a number.

## Syntax

**Hex** ( _number_ )
The required  _number_[argument](../../Glossary/vbe-glossary.md#argument) is any valid[numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) or [string expression](../../Glossary/vbe-glossary.md#string-expression) _._

## Remarks

If  _number_ is not already a whole number, it is rounded to the nearest whole number before being evaluated.


|**If  _number_ is**|**Hex returns**|
|:-----|:-----|
|[Null](../../Glossary/vbe-glossary.md#null)|Null|
|[Empty](../../Glossary/vbe-glossary.md#empty)|Zero (0)|
|Any other number|Up to eight hexadecimal characters|

You can represent hexadecimal numbers directly by preceding numbers in the proper range with  `&;H.` For example, For example, `&;H10` represents decimal 16 in hexadecimal notation.

## Example

This example uses the  **Hex** function to return the hexadecimal value of a number.


```vb
Dim MyHex
MyHex = Hex(5)    ' Returns 5.
MyHex = Hex(10)    ' Returns A.
MyHex = Hex(459)    ' Returns 1CB.
```


