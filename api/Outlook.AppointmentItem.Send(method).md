---
title: AppointmentItem.Send method (Outlook)
keywords: vbaol11.chm907
f1_keywords:
- vbaol11.chm907
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Send
ms.assetid: 72f2e997-55ef-b98b-fdd1-7f3b810a28ed
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.Send method (Outlook)

Sends the appointment.


## Syntax

_expression_. `Send`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

The  **Send** method sends an item using the default account specified for the session. In a session where multiple Microsoft Exchange accounts are defined in the profile, the first Exchange account added to the profile is the primary Exchange account, and is also the default account for the session. To specify a different account to send an item, set the **[SendUsingAccount](Outlook.AppointmentItem.SendUsingAccount.md)** property to the desired **[Account](Outlook.Account.md)** object and then call the **Send** method.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](Outlook.Application.CreateItem.md)** to create an appointment. The example sets the **[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** property to **olMeeting** to indicate the appointment as a meeting request, and sets a required attendee, an optional attendee, and a meeting location as a resource. The example then displays and sends the appointment item.


```vb
Sub CreateAppt() 
 Dim myItem As Object 
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 myItem.MeetingStatus = olMeeting 
 myItem.Subject = "Strategy Meeting" 
 myItem.Location = "Conf Rm All Stars" 
 myItem.Start = #9/24/2009 1:30:00 PM# 
 myItem.Duration = 90 
 Set myRequiredAttendee = myItem.Recipients.Add("Nate Sun") 
 myRequiredAttendee.Type = olRequired 
 Set myOptionalAttendee = myItem.Recipients.Add("Kevin Kennedy") 
 myOptionalAttendee.Type = olOptional 
 Set myResourceAttendee = myItem.Recipients.Add("Conf Rm All Stars") 
 myResourceAttendee.Type = olResource 
 myItem.Display 
 myItem.Send 
End Sub
```


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
