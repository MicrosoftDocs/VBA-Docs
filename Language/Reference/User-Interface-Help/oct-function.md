---
title: Oct function (Visual Basic for Applications)
keywords: vblr6.chm1008983
f1_keywords:
- vblr6.chm1008983
ms.prod: office
ms.assetid: 178a6099-9181-2160-2b97-e08c97f8b2bb
ms.date: 12/12/2018
localization_priority: Normal
---


# Oct function

Returns a **Variant** (**String**) representing the octal value of a number.

## Syntax

**Oct**(_number_)

The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) is any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) or [string expression](../../Glossary/vbe-glossary.md#string-expression).

## Remarks

If _number_ is not already a whole number, it is rounded to the nearest whole number before being evaluated.

|If _number_ is|Oct returns|
|:-----|:-----|
|[Null](../../Glossary/vbe-glossary.md#null)|**Null**|
|[Empty](../../Glossary/vbe-glossary.md#empty)|Zero (0)|
|Any other number|Up to 11 octal characters|

You can represent octal numbers directly by preceding numbers in the proper range with `&O`. For example, `&O10` is the octal notation for decimal 8.

## Example

This example uses the **Oct** function to return the octal value of a number.

```vb
Dim MyOct
MyOct = Oct(4)     ' Returns 4.
MyOct = Oct(8)    ' Returns 10.
MyOct = Oct(459)    ' Returns 713.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]