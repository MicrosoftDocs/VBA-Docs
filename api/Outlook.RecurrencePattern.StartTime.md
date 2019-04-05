---
title: RecurrencePattern.StartTime property (Outlook)
keywords: vbaol11.chm287
f1_keywords:
- vbaol11.chm287
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.StartTime
ms.assetid: 557e0f8d-c95d-e1f9-91a2-0734248d8628
ms.date: 06/08/2017
localization_priority: Normal
---


# RecurrencePattern.StartTime property (Outlook)

Returns or sets a  **Date** indicating the start time for a recurrence pattern. Read/write.


## Syntax

_expression_. `StartTime`

_expression_ A variable that represents a [RecurrencePattern](Outlook.RecurrencePattern.md) object.


## Remarks

This property is only valid for appointments.

When you create a  **[RecurrencePattern](Outlook.RecurrencePattern.md)** object and no time zones have been specified for the appointment, **StartTime** and **[EndTime](Outlook.RecurrencePattern.EndTime.md)** of the **RecurrencePattern** object are based on the time zone specified by **[Application.TimeZones.CurrentTimeZone](Outlook.TimeZones.CurrentTimeZone.md)**.

If you want to create a recurring appointment for a particular time zone, you should first create an  **[AppointmentItem](Outlook.AppointmentItem.md)**, set **[AppointmentItem.StartTimeZone](Outlook.AppointmentItem.StartTimeZone.md)**, and then call **[AppointmentItem.GetRecurrencePattern](Outlook.AppointmentItem.GetRecurrencePattern.md)**. The **RecurrencePattern** object returned will have both **StartTime** and **EndTime** based on the time zone specified by **AppointmentItem.StartTimeZone**. Note that in the **Appointment Recurrence** dialog box, the time indicated as **Start** is **RecurrencePattern.StartTime** which is based on **AppointmentItem.StartTimeZone**, but the time indicated as **End** is not always the same as **RecurrencePattern.EndTime** which is based on **AppointmentItem.StartTimeZone**; the displayed time value is based on **[AppointmentItem.EndTimeZone](Outlook.AppointmentItem.EndTimeZone.md)**.


## See also


[RecurrencePattern Object](Outlook.RecurrencePattern.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]