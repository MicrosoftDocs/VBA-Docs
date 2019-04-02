---
title: RecurrencePattern.RecurrenceType property (Outlook)
keywords: vbaol11.chm285
f1_keywords:
- vbaol11.chm285
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.RecurrenceType
ms.assetid: bc9b35b5-ef00-e5cf-09cc-ee8743efddcf
ms.date: 06/08/2017
localization_priority: Normal
---


# RecurrencePattern.RecurrenceType property (Outlook)

Returns or sets an  **[OlRecurrenceType](Outlook.OlRecurrenceType.md)** constant specifying the frequency of occurrences for the recurrence pattern. Read/write.


## Syntax

_expression_. `RecurrenceType`

_expression_ A variable that represents a [RecurrencePattern](Outlook.RecurrencePattern.md) object.


## Remarks

You must set the  **RecurrenceType** property before you set other properties for a **[RecurrencePattern](Outlook.RecurrencePattern.md)** object. The **RecurrencePattern** properties that you can set subsequently depends on the value of **RecurrenceType**, as shown in the following table:



| **OlRecurrenceType**| **Valid RecurrencePattern Properties**|
| **olRecursDaily**| **[Duration](Outlook.RecurrencePattern.Duration.md)**, **[EndTime](Outlook.RecurrencePattern.EndTime.md)**, **[Interval](Outlook.RecurrencePattern.Interval.md)**, **[NoEndDate](Outlook.RecurrencePattern.NoEndDate.md)**, **[Occurrences](Outlook.RecurrencePattern.Occurrences.md)**, **[PatternStartDate](Outlook.RecurrencePattern.PatternStartDate.md)**, **[PatternEndDate](Outlook.RecurrencePattern.PatternEndDate.md)**, **[StartTime](Outlook.RecurrencePattern.StartTime.md)**|
| **olRecursWeekly**| **[DayOfWeekMask](Outlook.RecurrencePattern.DayOfWeekMask.md)**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|
| **olRecursMonthly**| **[DayOfMonth](Outlook.RecurrencePattern.DayOfMonth.md)**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|
| **olRecursMonthNth**| **DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **[Instance](Outlook.RecurrencePattern.Instance.md)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|
| **olRecursYearly**| **DayOfMonth**, **Duration**, **EndTime**, **Interval**, **MonthOfYear**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|
| **olRecursYearNth**| **DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **Instance**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|

## Example

This Visual Basic for Applications example uses  **[GetRecurrencePattern](Outlook.AppointmentItem.GetRecurrencePattern.md)** to obtain the **[RecurrencePattern](Outlook.RecurrencePattern.md)** object for the newly-created **[AppointmentItem](Outlook.AppointmentItem.md)**. The properties, **RecurrenceType**, **DayOfWeekMask**, **[MonthOfYear](Outlook.RecurrencePattern.MonthOfYear.md)**, **Instance**, **Occurrences**, **StartTime**, **EndTime**, and **[Subject](Outlook.AppointmentItem.Subject.md)** are set, the appointment is saved and then displayed with the pattern: "Occurs the first Monday of June effective 6/1/2007 until 6/6/2016 from 2:00 PM to 5:00 PM."


```vb
Sub RecurringYearNth() 
 
 Dim oAppt As AppointmentItem 
 
 Dim oPattern As RecurrencePattern 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 Set oPattern = oAppt.GetRecurrencePattern 
 
 With oPattern 
 
 .RecurrenceType = olRecursYearNth 
 
 .DayOfWeekMask = olMonday 
 
 .MonthOfYear = 6 
 
 .Instance = 1 
 
 .Occurrences = 10 
 
 .Duration = 180 
 
 .PatternStartDate = #6/1/2007# 
 
 .StartTime = #2:00:00 PM# 
 
 .EndTime = #5:00:00 PM# 
 
 End With 
 
 oAppt.Subject = "Recurring YearNth Appointment" 
 
 oAppt.Save 
 
 oAppt.Display 
 
End Sub
```


## See also


[RecurrencePattern Object](Outlook.RecurrencePattern.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]