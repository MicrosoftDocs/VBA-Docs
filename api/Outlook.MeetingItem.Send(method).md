---
title: MeetingItem.Send method (Outlook)
keywords: vbaol11.chm1458
f1_keywords:
- vbaol11.chm1458
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Send
ms.assetid: d9a6ea8c-2146-06ec-aa8b-6e39fd60a916
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.Send method (Outlook)

Sends the meeting item.


## Syntax

_expression_. `Send`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

When you create a meeting request programmatically, you first create an  **[AppointmentItem](Outlook.AppointmentItem.md)** object instead of a **[MeetingItem](Outlook.MeetingItem.md)** object. To indicate the appointment is a meeting, set the **[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** property of the **AppointmentItem** object to **olMeeting**. To send the meeting request, apply the **[Send](Outlook.AppointmentItem.Send(method).md)** method to that **AppointmentItem** object.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)



[How to: Create an Appointment as a Meeting on the Calendar](../outlook/How-to/Items-Folders-and-Stores/create-an-appointment-as-a-meeting-on-the-calendar.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]