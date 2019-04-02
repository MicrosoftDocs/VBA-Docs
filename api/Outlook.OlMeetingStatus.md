---
title: OlMeetingStatus enumeration (Outlook)
keywords: vbaol11.chm3070
f1_keywords:
- vbaol11.chm3070
ms.prod: outlook
api_name:
- Outlook.OlMeetingStatus
ms.assetid: da83b8ed-267e-c055-13ce-11067e224e9d
ms.date: 06/08/2017
localization_priority: Normal
---


# OlMeetingStatus enumeration (Outlook)

Indicates the status of the meeting.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olMeeting**|1|The meeting has been scheduled.|
| **olMeetingCanceled**|5|The scheduled meeting has been cancelled.|
| **olMeetingReceived**|3|The meeting request has been received.|
| **olMeetingReceivedAndCanceled**|7|The scheduled meeting has been cancelled but still appears on the user's calendar.|
| **olNonMeeting**|0|An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar.|

## Remarks

See [AppointmentItem.MeetingStatus property (Outlook)](Outlook.AppointmentItem.MeetingStatus.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]