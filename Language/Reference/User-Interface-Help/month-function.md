---
title: Month function (Visual Basic for Applications)
keywords: vblr6.chm1008977
f1_keywords:
- vblr6.chm1008977
ms.prod: office
ms.assetid: d0b3cfef-d192-166d-2dcf-c60b08213327
ms.date: 12/13/2018
localization_priority: Normal
---


# Month function

Returns a **Variant** (**Integer**) specifying a whole number between 1 and 12, inclusive, representing the month of the year.

## Syntax

**Month**(_date_)

The required _date_ [argument](../../Glossary/vbe-glossary.md#argument) is any [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a date. If _date_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

> [!NOTE] 
> If the **[Calendar](calendar-property.md)** property setting is Gregorian, the returned integer represents the Gregorian day of the week for the date argument. If the calendar is Hijri, the returned integer represents the Hijri day of the week for the date argument. For Hijri dates, the argument number is any numeric expression that can represent a date and/or time from 1/1/100 (Gregorian Aug 2, 718) through 4/3/9666 (Gregorian Dec 31, 9999).


## Example

This example uses the **Month** function to obtain the month from a specified date. In the development environment, the date literal is displayed in short date format by using the locale settings of your code.

```vb
Dim MyDate, MyMonth
MyDate = #February 12, 1969#    ' Assign a date.
MyMonth = Month(MyDate)    ' MyMonth contains 2.
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
