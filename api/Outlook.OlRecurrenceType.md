---
title: OlRecurrenceType enumeration (Outlook)
keywords: vbaol11.chm3077
f1_keywords:
- vbaol11.chm3077
ms.prod: outlook
api_name:
- Outlook.OlRecurrenceType
ms.assetid: 63bc267e-6b9d-2cb5-3a96-4beb41afff72
ms.date: 06/08/2017
localization_priority: Normal
---


# OlRecurrenceType enumeration (Outlook)

Specifies the recurrence pattern type.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olRecursDaily**|0|Represents a daily recurrence pattern.|
| **olRecursMonthly**|2|Represents a monthly recurrence pattern.|
| **olRecursMonthNth**|3|Represents a MonthNth recurrence pattern. See  **[RecurrencePattern.Instance](Outlook.RecurrencePattern.Instance.md)** property.|
| **olRecursWeekly**|1|Represents a weekly recurrence pattern.|
| **olRecursYearly**|5|Represents a yearly recurrence pattern.|
| **olRecursYearNth**|6|Represents a YearNth recurrence pattern. See  **[RecurrencePattern.Instance](Outlook.RecurrencePattern.Instance.md)** property.|

## Remarks

Used by the [RecurrencePattern.RecurrenceType property (Outlook)](Outlook.RecurrencePattern.RecurrenceType.md) of an [AppointmentItem object (Outlook)](Outlook.AppointmentItem.md) to specify the frequency of occurrences of the appointment.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]