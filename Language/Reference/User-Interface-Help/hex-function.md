---
title: Hex function (Visual Basic for Applications)
keywords: vblr6.chm1011362
f1_keywords:
- vblr6.chm1011362
ms.prod: office
ms.assetid: 79a403a9-61af-0991-8f13-60c1033f158a
ms.date: 12/12/2018
localization_priority: Normal
---


# Hex function

Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) representing the hexadecimal value of a number.

## Syntax

**Hex**(_number_)

The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) is any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) or [string expression](../../Glossary/vbe-glossary.md#string-expression).

<br/>

|If _number_ is|Hex returns|
|:-----|:-----|
|-2,147,483,648 to 2,147,483,647|Up to eight hexadecimal characters|
|[Null](../../Glossary/vbe-glossary.md#null)|Null|
|[Empty](../../Glossary/vbe-glossary.md#empty)|Zero (0)|


## Remarks

If _number_ is not a whole number, it is rounded to the nearest whole number before being evaluated.

For the opposite of **Hex**, precede a hexadecimal value with **&H**. For example, `Hex(255)` returns the string FF and `&HFF` returns the number 255.


## Example

This example uses the **Hex** function to return the hexadecimal value of a number.

```vb
Dim MyHex
MyHex = Hex(5)    ' Returns 5.
MyHex = Hex(10)    ' Returns A.
MyHex = Hex(459)    ' Returns 1CB.
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
