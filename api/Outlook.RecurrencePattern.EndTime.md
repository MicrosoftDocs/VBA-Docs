---
title: RecurrencePattern.EndTime property (Outlook)
keywords: vbaol11.chm276
f1_keywords:
- vbaol11.chm276
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.EndTime
ms.assetid: 7babda13-9e57-4c80-1ab3-56025753ed9d
ms.date: 06/08/2017
localization_priority: Normal
---


# RecurrencePattern.EndTime property (Outlook)

Returns or sets a  **Date** indicating the end time for a recurrence pattern. Read/write.


## Syntax

_expression_. `EndTime`

_expression_ A variable that represents a [RecurrencePattern](Outlook.RecurrencePattern.md) object.


## Remarks

This property is only valid for appointments. 

When you create a  **[RecurrencePattern](Outlook.RecurrencePattern.md)** object and no time zones have been specified for the appointment, **[StartTime](Outlook.RecurrencePattern.StartTime.md)** and **EndTime** of the **RecurrencePattern** object are based on the time zone specified by **[Application.TimeZones.CurrentTimeZone](Outlook.TimeZones.CurrentTimeZone.md)**.

If you want to create a recurring appointment for a particular time zone, you should first create an  **[AppointmentItem](Outlook.AppointmentItem.md)**, set **[AppointmentItem.StartTimeZone](Outlook.AppointmentItem.StartTimeZone.md)**, and then call **[AppointmentItem.GetRecurrencePattern](Outlook.AppointmentItem.GetRecurrencePattern.md)**. The **RecurrencePattern** object returned will have both **StartTime** and **EndTime** based on the time zone specified by **AppointmentItem.StartTimeZone**. Note that in the **Appointment Recurrence** dialog box, the time indicated as **Start** is **RecurrencePattern.StartTime** which is based on **AppointmentItem.StartTimeZone**, but the time indicated as **End** is not always the same as **RecurrencePattern.EndTime** which is based on **AppointmentItem.StartTimeZone**; the displayed time value is based on **[AppointmentItem.EndTimeZone](Outlook.AppointmentItem.EndTimeZone.md)**.


## See also


[RecurrencePattern Object](Outlook.RecurrencePattern.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]