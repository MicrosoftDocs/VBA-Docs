---
title: RecurrencePattern object (Outlook)
keywords: vbaol11.chm268
f1_keywords:
- vbaol11.chm268
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern
ms.assetid: 36c098f7-59fb-879a-5173-ed0260d13fa4
ms.date: 06/08/2017
localization_priority: Normal
---


# RecurrencePattern object (Outlook)

Represents the pattern of incidence of recurring appointments and tasks for the associated  **[AppointmentItem](Outlook.AppointmentItem.md)** and **[TaskItem](Outlook.TaskItem.md)** object.


## Remarks

Use the  **GetRecurrencePattern** method to return the **RecurrencePattern** object associated with an **AppointmentItem** or **TaskItem** object.

Calling  **GetRecurrencePattern** or **ClearRecurrencePattern** has the side effect of setting the **IsRecurring** property of the item accordingly. This property can be used as required for efficient filtering of the **[Items](Outlook.Items.md)** object.

The type of recurrence pattern is indicated by the  **[RecurrenceType](Outlook.RecurrencePattern.RecurrenceType.md)** property. The **RecurrenceType** property is the first property you should set.

The following properties are valid for all recurrence patterns:  **[EndTime](Outlook.RecurrencePattern.EndTime.md)**, **[Occurrences](Outlook.RecurrencePattern.Occurrences.md)**, **StartDate**, **[StartTime](Outlook.RecurrencePattern.StartTime.md)**, or **Type**.

The following table shows the properties that are valid for the different recurrence types. An error occurs if the item is saved and the property is null or contains an invalid value. Monthly and yearly patterns are only valid for a single day. Weekly patterns are only valid as the  **Or** of the **[DayOfWeekMask](Outlook.RecurrencePattern.DayOfWeekMask.md)**.



|**RecurrenceType**|**Properties**|**Examples**|
|:-----|:-----|:-----|
|**olRecursDaily**|**[Duration](Outlook.RecurrencePattern.Duration.md)**, **EndTime**, **[Interval](Outlook.RecurrencePattern.Interval.md)**, **[NoEndDate](Outlook.RecurrencePattern.NoEndDate.md)**, **Occurrences**, **[PatternStartDate](Outlook.RecurrencePattern.PatternStartDate.md)**, **[PatternEndDate](Outlook.RecurrencePattern.PatternEndDate.md)**, **StartTime**|A value N for  **Interval** is every N days.|
|**olRecursWeekly**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N weeks. An example of **DayofWeekMask** is every Tuesday, Wednesday, and Thursday.|
|**olRecursMonthly**|**[DayOfMonth](Outlook.RecurrencePattern.DayOfMonth.md)**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N months. A value N for **DayofMonth** is every Nth day of the month.|
|**olRecursMonthNth**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **[Instance](Outlook.RecurrencePattern.Instance.md)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N months. An example of value N for **Instance** is every Nth Tuesday. An example of **DayofWeekMask** is every Tuesday and Wednesday.|
|**olRecursYearly**|**DayOfMonth**, **Duration**, **EndTime**, **Interval**, **[MonthOfYear](Outlook.RecurrencePattern.MonthOfYear.md)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **DayofMonth** is the Nth day of the month. An example of **MonthOfYear** is February.|
|**olRecursYearNth**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **Instance**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|An example of value N for  **Instance** is the Nth Tuesday. An example of **DayofWeekMask** is Tuesday, Wednesday, and Thursday. An example of **MonthOfYear** is February.|

When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](Outlook.Exception.md)** or **[RecurrencePattern](Outlook.RecurrencePattern.md)** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object. For a code example, see the topic for the **AppointmentItem** object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.


## Methods



|Name|
|:-----|
|[GetOccurrence](Outlook.RecurrencePattern.GetOccurrence.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.RecurrencePattern.Application.md)|
|[Class](Outlook.RecurrencePattern.Class.md)|
|[DayOfMonth](Outlook.RecurrencePattern.DayOfMonth.md)|
|[DayOfWeekMask](Outlook.RecurrencePattern.DayOfWeekMask.md)|
|[Duration](Outlook.RecurrencePattern.Duration.md)|
|[EndTime](Outlook.RecurrencePattern.EndTime.md)|
|[Exceptions](Outlook.RecurrencePattern.Exceptions.md)|
|[Instance](Outlook.RecurrencePattern.Instance.md)|
|[Interval](Outlook.RecurrencePattern.Interval.md)|
|[MonthOfYear](Outlook.RecurrencePattern.MonthOfYear.md)|
|[NoEndDate](Outlook.RecurrencePattern.NoEndDate.md)|
|[Occurrences](Outlook.RecurrencePattern.Occurrences.md)|
|[Parent](Outlook.RecurrencePattern.Parent.md)|
|[PatternEndDate](Outlook.RecurrencePattern.PatternEndDate.md)|
|[PatternStartDate](Outlook.RecurrencePattern.PatternStartDate.md)|
|[RecurrenceType](Outlook.RecurrencePattern.RecurrenceType.md)|
|[Regenerate](Outlook.RecurrencePattern.Regenerate.md)|
|[Session](Outlook.RecurrencePattern.Session.md)|
|[StartTime](Outlook.RecurrencePattern.StartTime.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[RecurrencePattern Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]