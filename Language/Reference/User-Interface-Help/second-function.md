---
title: Second function (Visual Basic for Applications)
keywords: vblr6.chm1009011
f1_keywords:
- vblr6.chm1009011
ms.prod: office
ms.assetid: fef87486-ccda-23e7-04a5-5e484ce66543
ms.date: 12/13/2018
localization_priority: Normal
---


# Second function

Returns a **Variant** (**Integer**) specifying a whole number between 0 and 59, inclusive, representing the second of the minute.

## Syntax

**Second**(_time_)

The required _time_ [argument](../../Glossary/vbe-glossary.md#argument) is any [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a time. If  _time_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

## Example

This example uses the **Second** function to obtain the second of the minute from a specified time. In the development environment, the time literal is displayed in short time format by using the locale settings of your code.

```vb
Dim MyTime, MySecond
MyTime = #4:35:17 PM#    ' Assign a time.
MySecond = Second(MyTime)    ' MySecond contains 17.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]