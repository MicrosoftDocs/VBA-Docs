---
title: AppointmentItem.Duration property (Outlook)
keywords: vbaol11.chm878
f1_keywords:
- vbaol11.chm878
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Duration
ms.assetid: eea64bdd-c19b-01c7-4fdb-111df86de2c4
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.Duration property (Outlook)

Returns or sets a  **Long** indicating the duration (in minutes) of the **[AppointmentItem](Outlook.AppointmentItem.md)**. Read/write.


## Syntax

_expression_. `Duration`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Example

This Visual Basic for Applications example uses  **[Application.CreateItem](Outlook.Application.CreateItem.md)** to create an appointment and uses **[AppointmentItem.MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** to set the meeting status to "Meeting" to turn it into a meeting request with both a required and an optional attendee.


```vb
Sub ScheduleMeeting() 
 
 Dim myItem as AppointmentItem 
 
 Dim myRequiredAttendee As Recipient 
 
 Dim myOptionalAttendee As Recipient 
 
 Dim myResourceAttendee As Recipient 
 
 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 
 myItem.MeetingStatus = olMeeting 
 
 myItem.Subject = "Strategy Meeting" 
 
 myItem.Location = "Conference Room B" 
 
 myItem.Start = #9/24/2002 1:30:00 PM# 
 
 myItem.Duration = 90 
 
 Set myRequiredAttendee = myItem.Recipients.Add ("Nate Sun") 
 
 myRequiredAttendee.Type = olRequired 
 
 Set myOptionalAttendee = myItem.Recipients.Add ("Kevin Kennedy") 
 
 myOptionalAttendee.Type = olOptional 
 
 Set myResourceAttendee = myItem.Recipients.Add("Conference Room B") 
 
 myResourceAttendee.Type = olResource 
 
 myItem.Display 
 
End Sub
```


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]