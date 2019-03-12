---
title: Weekday function (Visual Basic for Applications)
keywords: vblr6.chm1009058
f1_keywords:
- vblr6.chm1009058
ms.prod: office
ms.assetid: 4e6197a7-5c55-e5cd-5164-ce1d31a9f80c
ms.date: 12/13/2018
localization_priority: Normal
---


# Weekday function

Returns a **Variant** (**Integer**) containing a whole number representing the day of the week.

## Syntax

**Weekday**(_date_, [ _firstdayofweek_ ])

<br/>

The **Weekday** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_date_|Required. [Variant](../../Glossary/vbe-glossary.md#variant-data-type), [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a date. If _date_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.|
|_firstdayofweek_|Optional. A [constant](../../Glossary/vbe-glossary.md#constant) that specifies the first day of the week. If not specified, **vbSunday** is assumed.|

## Settings

The _firstdayofweek_ argument has these settings:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbUseSystem**|0|Use the NLS API setting.|
|**vbSunday**|1|Sunday (default)|
|**vbMonday**|2|Monday|
|**vbTuesday**|3|Tuesday|
|**vbWednesday**|4|Wednesday|
|**vbThursday**|5|Thursday|
|**vbFriday**|6|Friday|
|**vbSaturday**|7|Saturday|

## Return values

The **Weekday** function can return any of these values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbSunday**|1|Sunday|
|**vbMonday**|2|Monday|
|**vbTuesday**|3|Tuesday|
|**vbWednesday**|4|Wednesday|
|**vbThursday**|5|Thursday|
|**vbFriday**|6|Friday|
|**vbSaturday**|7|Saturday|

## Remarks

If the **[Calendar](calendar-property.md)** property setting is Gregorian, the returned integer represents the Gregorian day of the week for the date argument. 

If the calendar is Hijri, the returned integer represents the Hijri day of the week for the date argument. For Hijri dates, the argument number is any numeric expression that can represent a date and/or time from 1/1/100 (Gregorian Aug 2, 718) through 4/3/9666 (Gregorian Dec 31, 9999).

## Example

This example uses the **Weekday** function to obtain the day of the week from a specified date.

```vb
Dim MyDate, MyWeekDay
MyDate = #February 12, 1969#    ' Assign a date.
MyWeekDay = Weekday(MyDate)    ' MyWeekDay contains 4 because 
    ' MyDate represents a Wednesday.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
