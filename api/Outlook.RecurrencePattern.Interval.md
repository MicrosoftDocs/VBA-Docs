---
title: RecurrencePattern.Interval property (Outlook)
keywords: vbaol11.chm279
f1_keywords:
- vbaol11.chm279
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.Interval
ms.assetid: e3220174-38dc-d1e3-8d26-b3f208b554a4
ms.date: 06/08/2017
localization_priority: Normal
---


# RecurrencePattern.Interval property (Outlook)

Returns or sets a **Long** specifying the number of units of a given recurrence type between occurrences. Read/write.


## Syntax

_expression_. `Interval`

_expression_ A variable that represents a [RecurrencePattern](Outlook.RecurrencePattern.md) object.


## Remarks

The **Interval** property must be set before setting **[PatternEndDate](Outlook.RecurrencePattern.PatternEndDate.md)**.

For example, setting the  **Interval** property to 2 and the **[RecurrenceType](Outlook.RecurrencePattern.RecurrenceType.md)** property to **olRecursWeekly** would cause the pattern to occur every second week.

When  **RecurrenceType** is set to **olRecursYearNth** or **olRecursYear**, the **Interval** property indicates the number of years between occurrences. For example, **Interval** equals 1 indicates the recurrence is every year, **Interval** equals 2 indicates the recurrence is every 2 years, and so on.


## Example

This Visual Basic for Applications example uses  **[GetRecurrencePattern](Outlook.AppointmentItem.GetRecurrencePattern.md)** to obtain the **[RecurrencePattern](Outlook.RecurrencePattern.md)** object for the newly-created **[AppointmentItem](Outlook.AppointmentItem.md)**. The properties, **[RecurrenceType](Outlook.RecurrencePattern.RecurrenceType.md)**, **[DayOfWeekMask](Outlook.RecurrencePattern.DayOfWeekMask.md)**, **[PatternStartDate](Outlook.RecurrencePattern.PatternStartDate.md)**, **[Interval](Outlook.RecurrencePattern.Interval.md)**, **[PatternEndDate](Outlook.RecurrencePattern.PatternEndDate.md)**, and **[Subject](Outlook.AppointmentItem.Subject.md)** are set, the appointment is saved and then displayed with the pattern: "Occurs every 3 week(s) on Monday effective 1/21/2003 until 12/21/2004 from 2:00 PM to 5:00 PM."


```vb
Sub CreateAppointment() 
 
 Dim myApptItem As AppointmentItem 
 
 Dim myRecurrPatt As RecurrencePattern 
 
 
 
 
 
 Set myApptItem = Application.CreateItem(olAppointmentItem) 
 
 Set myRecurrPatt = myApptItem.GetRecurrencePattern 
 
 myRecurrPatt.RecurrenceType = olRecursWeekly 
 
 myRecurrPatt.DayOfWeekMask = olMonday 
 
 myRecurrPatt.PatternStartDate = #1/21/2003 2:00:00 PM# 
 
 myRecurrPatt.Interval = 3 
 
 myRecurrPatt.PatternEndDate = #12/21/2004 5:00:00 PM# 
 
 myApptItem.Subject = "Important Appointment" 
 
 myApptItem.Save 
 
 myApptItem.Display 
 
 Set myOlApp = Nothing 
 
 Set myApptItem = Nothing 
 
 Set myRecurrPatt = Nothing 
 
End Sub
```


## See also


[RecurrencePattern Object](Outlook.RecurrencePattern.md)



[How to: Create an Appointment as a Meeting on the Calendar](../outlook/How-to/Items-Folders-and-Stores/create-an-appointment-as-a-meeting-on-the-calendar.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]