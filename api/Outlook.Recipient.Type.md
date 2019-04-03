---
title: Recipient.Type property (Outlook)
keywords: vbaol11.chm2355
f1_keywords:
- vbaol11.chm2355
ms.prod: outlook
api_name:
- Outlook.Recipient.Type
ms.assetid: 3bdc616c-f008-ec95-0a92-0f704eedee34
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient.Type property (Outlook)

Returns or sets a **Long** representing the type of recipient. Read/write.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a [Recipient](Outlook.Recipient.md) object.


## Remarks

Depending on the type of recipient, this property returns or sets a **Long** corresponding to the numeric equivalent of one of the following constants:


- **[JournalItem](Outlook.JournalItem.md)** recipient: the **[OlJournalRecipientType](Outlook.OlJournalRecipientType.md)** constant **olAssociatedContact**.
    
- **[MailItem](Outlook.MailItem.md)** recipient: one of the following **[OlMailRecipientType](Outlook.OlMailRecipientType.md)** constants: **olBCC**, **olCC**, **olOriginator**, or **olTo**.
    
- **[MeetingItem](Outlook.MeetingItem.md)** recipient: one of the following **[OlMeetingRecipientType](Outlook.OlMeetingRecipientType.md)** constants: **olOptional**, **olOrganizer**, **olRequired**, or **olResource**.
    
- **[TaskItem](Outlook.TaskItem.md)** recipient: either of the following **[OlTaskRecipientType](Outlook.OlTaskRecipientType.md)** constants: **olFinalStatus**, or **olUpdate**.
    


This property may not always return the appropriate recipient type for a conference room. For instance, a conference room may be specified as a required recipient in a meeting request, in which case this property will not return **olResource** for that conference room.

To reliably determine if a recipient is a conference room, use the Messaging API (MAPI) property, **[PidTagDisplayTypeEx](overview/Outlook.md)**, of the **[Recipient](Outlook.Recipient.md)** object. You can access this property using the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object in the Outlook object model. The **PidTagDisplayTypeEx** property is represented as "http://schemas.microsoft.com/mapi/proptag/0x39050003" in the MAPI proptag namespace. Note that the **PidTagDisplayTypeEx** property is not available in versions of Microsoft Exchange Server earlier than Microsoft Exchange Server 2007; in such earlier versions of Exchange Server, you can use the **Recipient.Type** property and assume that a recipient having a type other than **olResource** is not a conference room.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the **PropertyAccessor** on the **PidTagDisplayTypeEx** property for each of the **Recipient** objects in the **[Recipients](Outlook.Recipients.md)** collection of a meeting request. If the value of that property is 7 (the value for the MAPI constant **DT_ROOM** as defined in the MAPI header file, mapidefs.h), then that recipient is a conference room. This example assumes that there is a meeting request in the current inspector.


```vb
Sub DemoMeetingRecipients() 
 Dim myAppointment As Outlook.AppointmentItem 
 Dim myPA As Outlook.PropertyAccessor 
 Dim d As Long 
 Dim myInt As Long 
 
 Set myAppointment = Application.ActiveInspector.CurrentItem 
 
 For d = 1 To myAppointment.Recipients.count 
 Debug.Print myAppointment.Recipients.item(d).name 
 Debug.Print myAppointment.Recipients.item(d).Type 
 Set myPA = myAppointment.Recipients.item(d).PropertyAccessor 
 myInt = myPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39050003") 
 Debug.Print myInt 
 Debug.Print "---" 
 Next d 
End Sub
```

The following VBA example uses **[CreateItem](Outlook.Application.CreateItem.md)** to create an appointment and uses **[MeetingStatus](Outlook.AppointmentItem.MeetingStatus.md)** to set the meeting status to "Meeting" to turn it into a meeting request with both a required and an optional attendee. The recipient names should be replaced with valid names to avoid errors.




```vb
Sub ScheduleMeeting() 
 Dim myItem as Outlook.AppointmentItem 
 Dim myRequiredAttendee As Outlook.Recipient 
 Dim myOptionalAttendee As Outlook.Recipient 
 Dim myResourceAttendee As Outlook.Recipient 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 myItem.MeetingStatus = olMeeting 
 myItem.Subject = "Strategy Meeting" 
 myItem.Location = "Conference Room B" 
 myItem.Start = #9/24/2003 1:30:00 PM# 
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


[Recipient Object](Outlook.Recipient.md)




[Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
