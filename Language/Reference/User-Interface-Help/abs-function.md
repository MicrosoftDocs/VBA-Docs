---
title: Abs function (Visual Basic for Applications)
keywords: vblr6.chm1008850
f1_keywords:
- vblr6.chm1008850
ms.prod: office
ms.assetid: b2184f54-bf2b-a3da-f1c8-b38575a213eb
ms.date: 12/11/2018
localization_priority: Normal
---


# Abs function

Returns a value of the same type that is passed to it specifying the absolute value of a number.

## Syntax

**Abs**(_number_)
 
The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) can be any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression). If _number_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned; if it is an uninitialized [variable](../../Glossary/vbe-glossary.md#variable), zero is returned.

## Remarks

The absolute value of a number is its unsigned magnitude. For example, `ABS(-1)` and `ABS(1)` both return `1`.

## Example

This example uses the **Abs** function to compute the absolute value of a number.


```vb
Dim MyNumber
MyNumber = Abs(50.3)    ' Returns 50.3.
MyNumber = Abs(-50.3)    ' Returns 50.3.
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
