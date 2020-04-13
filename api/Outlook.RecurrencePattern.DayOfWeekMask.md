---
title: RecurrencePattern.DayOfWeekMask property (Outlook)
keywords: vbaol11.chm274
f1_keywords:
- vbaol11.chm274
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.DayOfWeekMask
ms.assetid: 79268798-90ab-4161-5a6e-97669daa475a
ms.date: 06/08/2017
localization_priority: Normal
---


# RecurrencePattern.DayOfWeekMask property (Outlook)

Returns or sets an **[OlDaysOfWeek](Outlook.OlDaysOfWeek.md)** constant representing the mask for the days of the week on which the recurring appointment or task occurs. Read/write.


## Syntax

_expression_. `DayOfWeekMask`

_expression_ A variable that represents a [RecurrencePattern](Outlook.RecurrencePattern.md) object.


## Remarks

The **DayOfWeekMask** should be set after the **[RecurrenceType](Outlook.RecurrencePattern.RecurrenceType.md)** property has been set and before the **[PatternEndDate](Outlook.RecurrencePattern.PatternEndDate.md)** and **[PatternStartDate](Outlook.RecurrencePattern.PatternStartDate.md)** properties are set.

Monthly and yearly patterns are only valid for a single day. Weekly patterns are only valid as the  **Or** of the **DayOfWeekMask**.


## Example

This Visual Basic for Applications example uses  **[GetRecurrencePattern](Outlook.AppointmentItem.GetRecurrencePattern.md)** to obtain the **[RecurrencePattern](Outlook.RecurrencePattern.md)** object for the newly-created **[AppointmentItem](Outlook.AppointmentItem.md)**. The properties, **[RecurrenceType](Outlook.RecurrencePattern.RecurrenceType.md)**, **DayOfWeekMask**, **PatternStartDate**, **PatternEndDate**, **[Duration](Outlook.RecurrencePattern.Duration.md)**, **[StartTime](Outlook.RecurrencePattern.StartTime.md)**, **[EndTime](Outlook.RecurrencePattern.EndTime.md)**, and **[Subject](Outlook.AppointmentItem.Subject.md)** are set, the appointment is saved and then displayed with the pattern: "Occurs every Monday, Wednesday, and Friday effective 7/10/2006 until 8/25/2006 from 2:00 PM to 3:00 PM."


```vb
Sub RecurringAppointmentEveryMondayWednesdayFriday() 
 
 Dim oAppt As AppointmentItem 
 
 Dim oPattern As RecurrencePattern 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 Set oPattern = oAppt.GetRecurrencePattern 
 
 With oPattern 
 
 .RecurrenceType = olRecursWeekly 
 
 .DayOfWeekMask = olMonday Or olWednesday Or olFriday 
 
 .PatternStartDate = #7/10/2006# 
 
 .PatternEndDate = #8/25/2006# 
 
 .Duration = 60 
 
 .StartTime = #2:00:00 PM# 
 
 .EndTime = #3:00:00 PM# 
 
 End With 
 
 oAppt.Subject = "Recurring Appointment Monday Wednesday Friday" 
 
 oAppt.Save 
 
 oAppt.Display 
 
End Sub
```

Similar to the last example, this Visual Basic for Applications example also uses  **GetRecurrencePattern** to obtain the **RecurrencePattern** object for the newly-created **AppointmentItem**. The properties, **RecurrenceType**, **DayOfWeekMask**, **PatternStartDate**, **PatternEndDate**, **Duration**, **StartTime**, **EndTime**, and **Subject** are set, the appointment is saved and then displayed with the pattern: "Occurs every Monday, Tuesday, Wednesday, Thursday, and Friday effective 7/10/2006 until 8/4/2006."




```vb
Sub RecurringEventEveryWeekday() 
 
 Dim oPattern As Outlook.RecurrencePattern 
 
 Dim oAppt As Outlook.AppointmentItem 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 Set oPattern = oAppt.GetRecurrencePattern 
 
 With oPattern 
 
 .RecurrenceType = olRecursWeekly 
 
 .DayOfWeekMask = olMonday Or olTuesday Or olWednesday Or olThursday Or olFriday 
 
 .PatternStartDate = #7/10/2006# 
 
 .PatternEndDate = #8/4/2006# 
 
 .Duration = 1440 'Duration in minutes, for all day event = 24 * 60 
 
 .StartTime = #12:00:00 AM# 
 
 .EndTime = #12:00:00 AM# 
 
 End With 
 
 oAppt.Subject = "Recurring Event Every Weekday" 
 
 oAppt.Save 
 
 oAppt.Display 
 
End Sub
```


## See also


[RecurrencePattern Object](Outlook.RecurrencePattern.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]