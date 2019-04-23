---
title: Day function (Visual Basic for Applications)
keywords: vblr6.chm1008890
f1_keywords:
- vblr6.chm1008890
ms.prod: office
ms.assetid: 8d4d0b63-28d9-c6a3-bd96-3688b0f93a12
ms.date: 12/12/2018
localization_priority: Normal
---


# Day function

Returns a **Variant** (**Integer**) specifying a whole number between 1 and 31, inclusive, representing the day of the month.

## Syntax

**Day**(_date_)

The required _date_ [argument](../../Glossary/vbe-glossary.md#argument) is any [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a date. If _date_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

> [!NOTE] 
> If the **[Calendar](calendar-property.md)** property setting is Gregorian, the returned integer represents the Gregorian day of the month for the date argument. If the calendar is Hijri, the returned integer represents the Hijri day of the month for the date argument.


## Example

This example uses the **Day** function to obtain the day of the month from a specified date. In the development environment, the date literal is displayed in short format using the locale settings of your code.


```vb
Dim MyDate, MyDay
MyDate = #February 12, 1969#    ' Assign a date.
MyDay = Day(MyDate)    ' MyDay contains 12.


```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
