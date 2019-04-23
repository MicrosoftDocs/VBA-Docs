---
title: Minute function (Visual Basic for Applications)
keywords: vblr6.chm1008974
f1_keywords:
- vblr6.chm1008974
ms.prod: office
ms.assetid: 47b5924f-41cf-9c25-68df-3932f0d77f8b
ms.date: 12/13/2018
localization_priority: Normal
---


# Minute function

Returns a **Variant** (**Integer**) specifying a whole number between 0 and 59, inclusive, representing the minute of the hour.

## Syntax

**Minute**(_time_)

The required _time_ [argument](../../Glossary/vbe-glossary.md#argument) is any [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a time. If _time_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

## Example

This example uses the **Minute** function to obtain the minute of the hour from a specified time. In the development environment, the time literal is displayed in short time format by using the locale settings of your code.


```vb
Dim MyTime, MyMinute
MyTime = #4:35:17 PM#    ' Assign a time.
MyMinute = Minute(MyTime)    ' MyMinute contains 35.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]