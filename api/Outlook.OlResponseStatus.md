---
title: OlResponseStatus enumeration (Outlook)
keywords: vbaol11.chm3079
f1_keywords:
- vbaol11.chm3079
ms.prod: outlook
api_name:
- Outlook.OlResponseStatus
ms.assetid: b473d57a-76a1-0862-fecb-baf1cf317772
ms.date: 06/08/2017
localization_priority: Normal
---


# OlResponseStatus enumeration (Outlook)

Indicates the response to a meeting request.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olResponseAccepted**|3|Meeting accepted.|
| **olResponseDeclined**|4|Meeting declined.|
| **olResponseNone**|0|The appointment is a simple appointment and does not require a response.|
| **olResponseNotResponded**|5|Recipient has not responded.|
| **olResponseOrganized**|1|The  **AppointmentItem** is on the Organizer's calendar or the recipient is the **Organizer** of the meeting.|
| **olResponseTentative**|2|Meeting tentatively accepted.|

## Remarks

Used by [Recipient.MeetingResponseStatus property (Outlook)](Outlook.Recipient.MeetingResponseStatus.md) and [AppointmentItem.ResponseStatus property (Outlook)](Outlook.AppointmentItem.ResponseStatus.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]