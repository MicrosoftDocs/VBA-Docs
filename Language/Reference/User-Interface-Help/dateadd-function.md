---
title: DateAdd function (Visual Basic for Applications)
keywords: vblr6.chm1013094
f1_keywords:
- vblr6.chm1013094
ms.prod: office
ms.assetid: 68d4e339-67b2-37e7-214d-318edd683b23
ms.date: 12/12/2018
localization_priority: Priority
---

# DateAdd function

Returns a **Variant** (**Date**) containing a date to which a specified time interval has been added.

## Syntax

**DateAdd**(_interval, number, date_)

The **DateAdd** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_interval_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) that is the interval of time you want to add.|
|_number_|Required. [Numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) that is the number of intervals you want to add. It can be positive (to get dates in the future) or negative (to get dates in the past).|
|_date_|Required. **Variant** (**Date**) or literal representing the date to which the interval is added.|

## Settings

The _interval_ [argument](../../Glossary/vbe-glossary.md#argument) has these settings:

<br/>

|Setting|Description|
|:-----|:-----|
|yyyy|Year|
|q|Quarter|
|m|Month|
|y|Day of year|
|d|Day|
|w|Weekday|
|ww|Week|
|h|Hour|
|n|Minute|
|s|Second|

## Remarks

You can use the **DateAdd** function to add or subtract a specified time interval from a date. For example, you can use **DateAdd** to calculate a date 30 days from today or a time 45 minutes from now.

To add days to _date_, you can use Day of Year ("y"), Day ("d"), or Weekday ("w").

> [!NOTE] 
> When you use the "w" interval (which includes all the days of the week, Sunday through Saturday) to add days to a date, the **DateAdd** function adds the total number of days that you specified to the date, instead of adding just the number of workdays (Monday through Friday) to the date, as you might expect.

The **DateAdd** function won't return an invalid date. The following example adds one month to January 31:

```vb
DateAdd("m", 1, "31-Jan-95")

```

In this case, **DateAdd** returns 28-Feb-95, not 31-Feb-95. If _date_ is 31-Jan-96, it returns 29-Feb-96 because 1996 is a leap year.

If the calculated date would precede the year 100 (that is, you subtract more years than are in _date_), an error occurs.

If _number_ isn't a [Long](../../Glossary/vbe-glossary.md#long-data-type) value, it is rounded to the nearest whole number before being evaluated.

> [!NOTE] 
> The format of the return value for **DateAdd** is determined by **Control Panel** settings, not by the format that is passed in the _date_ argument.

> [!NOTE] 
> For _date_, if the **[Calendar](calendar-property.md)** property setting is Gregorian, the supplied date must be Gregorian. If the calendar is Hijri, the supplied date must be Hijri. If month values are names, the name must be consistent with the current **Calendar** property setting. To minimize the possibility of month names conflicting with the current **Calendar** property setting, enter numeric month values (Short Date format).

## Example

This example takes a date and, using the **DateAdd** function, displays a corresponding date a specified number of months in the future.

```vb
Dim FirstDate As Date    ' Declare variables.
Dim IntervalType As String
Dim Number As Integer
Dim Msg As String
IntervalType = "m"    ' "m" specifies months as interval.
FirstDate = InputBox("Enter a date")
Number = InputBox("Enter number of months to add")
Msg = "New date: " & DateAdd(IntervalType, Number, FirstDate)
MsgBox Msg

```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
