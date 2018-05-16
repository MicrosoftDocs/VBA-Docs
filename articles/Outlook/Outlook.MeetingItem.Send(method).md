---
title: MeetingItem.Send Method (Outlook)
keywords: vbaol11.chm1458
f1_keywords:
- vbaol11.chm1458
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Send
ms.assetid: d9a6ea8c-2146-06ec-aa8b-6e39fd60a916
ms.date: 06/08/2017
---


# MeetingItem.Send Method (Outlook)

Sends the meeting item.


## Syntax

 _expression_ . **Send**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

When you create a meeting request programmatically, you first create an  **[AppointmentItem](Outlook.AppointmentItem.md)** object instead of a **[MeetingItem](Outlook.MeetingItem.md)** object. To indicate the appointment is a meeting, set the **[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** property of the **AppointmentItem** object to **olMeeting** . To send the meeting request, apply the **[Send](Outlook.AppointmentItem.Send(method).md)** method to that **AppointmentItem** object.


## See also


#### Concepts


[MeetingItem Object](Outlook.MeetingItem.md)
#### Other resources


[How to: Create an Appointment as a Meeting on the Calendar](http://msdn.microsoft.com/library/130b6ae1-d1a4-3805-7e9c-75543b93fff5%28Office.15%29.aspx)


