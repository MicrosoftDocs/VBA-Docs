---
title: AppointmentItem.Location property (Outlook)
keywords: vbaol11.chm882
f1_keywords:
- vbaol11.chm882
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Location
ms.assetid: bde4d455-15de-bb29-c27e-99c34836bd46
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.Location property (Outlook)

Returns or sets a **String** representing the specific office location (for example, Building 1 Room 1 or Suite 123) for the appointment. Read/write.


## Syntax

_expression_.**Location** 

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagOfficeLocation**.


## Example

This Visual Basic for Applications example uses  **[CreateItem](Outlook.Application.CreateItem.md)** to create an appointment and uses **[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** to set the meeting status to "Meeting" to turn it into a meeting request with both a required and an optional attendee.


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



[How to: Import Appointment XML Data into Outlook Appointment Objects](../outlook/How-to/Items-Folders-and-Stores/import-appointment-xml-data-into-outlook-appointment-objects-outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]