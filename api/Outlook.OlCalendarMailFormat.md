---
title: OlCalendarMailFormat enumeration (Outlook)
keywords: vbaol11.chm3117
f1_keywords:
- vbaol11.chm3117
ms.prod: outlook
api_name:
- Outlook.OlCalendarMailFormat
ms.assetid: b4b77080-1c8b-cfa4-3b3a-e59fec698bb1
ms.date: 06/08/2017
localization_priority: Normal
---


# OlCalendarMailFormat enumeration (Outlook)

Determines the format of the calendar information in the body of the  **[MailItem](Outlook.MailItem.md)** created by the **[ForwardAsICal](Outlook.CalendarSharing.ForwardAsICal.md)** method.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olCalendarMailFormatDailySchedule**|0|The calendar information is formatted as a daily schedule of appointments, containing an hour-by-hour breakdown of the calendar, showing both free and busy time blocks along with working-hour information. This layout is intended to help show recipients which times you are available. |
| **olCalendarMailFormatEventList**|1|The calendar information is formatted as a list of events, containing a list of the calendar appointments without showing any time blocks. This layout is intended to help show recipients the events scheduled for a given time period.|

## Remarks

For more information, see [Sharing Calendars](../outlook/How-to/Sharing/sharing-calendars.md) and [Export a Calendar using Payload Sharing](../outlook/How-to/Sharing/export-a-calendar-using-payload-sharing.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]