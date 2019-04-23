---
title: TimeSerial function (Visual Basic for Applications)
keywords: vblr6.chm1009044
f1_keywords:
- vblr6.chm1009044
ms.prod: office
ms.assetid: 5b08df07-bffb-ba69-7336-53067775fbf5
ms.date: 12/13/2018
localization_priority: Normal
---


# TimeSerial function

Returns a **Variant** (**Date**) containing the time for a specific hour, minute, and second.

## Syntax

**TimeSerial**(_hour_, _minute_, _second_)

<br/>

The **TimeSerial** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_hour_|Required; **Variant** (**Integer**). Number between 0 (12:00 A.M.) and 23 (11:00 P.M.), inclusive, or a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).|
|_minute_|Required; **Variant** (**Integer**). Any numeric expression.|
|_second_|Required; **Variant** (**Integer**). Any numeric expression.|

## Remarks

To specify a time, such as 11:59:59, the range of numbers for each **TimeSerial** argument should be in the normal range for the unit; that is, 0&ndash;23 for hours and 0&ndash;59 for minutes and seconds. However, you can also specify relative times for each [argument](../../Glossary/vbe-glossary.md#argument) by using any numeric expression that represents some number of hours, minutes, or seconds before or after a certain time. 

The following example uses [expressions](../../Glossary/vbe-glossary.md#expression) instead of absolute time numbers. The **TimeSerial** function returns a time for 15 minutes before (`-15`) six hours before noon (`12 - 6`), or 5:45:00 A.M.

```vb
TimeSerial(12 - 6, -15, 0)
```

When any argument exceeds the normal range for that argument, it increments to the next larger unit as appropriate. For example, if you specify 75 minutes, it is evaluated as one hour and 15 minutes. If any single argument is outside the range -32,768 to 32,767, an error occurs. If the time specified by the three arguments causes the date to fall outside the acceptable range of dates, an error occurs.

## Example

This example uses the **TimeSerial** function to return a time for the specified hour, minute, and second.

```vb
Dim MyTime
MyTime = TimeSerial(16, 35, 17)    ' MyTime contains serial 
    ' representation of 4:35:17 PM.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
