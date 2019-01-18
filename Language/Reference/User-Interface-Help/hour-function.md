---
title: Hour function (Visual Basic for Applications)
keywords: vblr6.chm1008939
f1_keywords:
- vblr6.chm1008939
ms.prod: office
ms.assetid: cf0800d1-6e26-71ad-ec8d-09e4876bf469
ms.date: 12/12/2018
localization_priority: Normal
---


# Hour function

Returns a **Variant** (**Integer**) specifying a whole number between 0 and 23, inclusive, representing the hour of the day.

## Syntax

**Hour**(_time_)

The required _time_ [argument](../../Glossary/vbe-glossary.md#argument) is any [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a time. If _time_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

## Example

This example uses the **Hour** function to obtain the hour from a specified time. In the development environment, the time literal is displayed in short time format by using the locale settings of your code.

```vb
Dim MyTime, MyHour
MyTime = #4:35:17 PM#    ' Assign a time.
MyHour = Hour(MyTime)    ' MyHour contains 16.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]