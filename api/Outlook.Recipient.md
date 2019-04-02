---
title: Recipient object (Outlook)
keywords: vbaol11.chm2339
f1_keywords:
- vbaol11.chm2339
ms.prod: outlook
api_name:
- Outlook.Recipient
ms.assetid: 8cee4d79-ec55-52a4-710b-6456944ca86d
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient object (Outlook)

Represents a user or resource in Outlook, generally a mail or mobile message addressee.

## Remarks

Use the **[Recipients](Outlook.Recipients.Item.md)** (_index_) method, where _index_ is the name or index number, to return a single **Recipient** object. The name can be a string that represents the display name, the alias, the full SMTP email address, or the mobile phone number of the recipient. A good practice is to use the SMTP email address for a mail message, and the mobile phone number for a mobile message.

Use the **[Add](Outlook.Recipients.Add.md)** method to create a new **Recipient** object and add it to the **[Recipients](Outlook.Recipients.md)** object. 

The **[Type](Outlook.Recipient.Type.md)** property of a new **Recipient** object is set to the default value for the associated **[AppointmentItem](Outlook.AppointmentItem.md)**, **[JournalItem](Outlook.JournalItem.md)**, **[MailItem](Outlook.MailItem.md)**, **[MeetingItem](Outlook.MeetingItem.md)**, or **[TaskItem](Outlook.TaskItem.md)** object and must be reset to indicate another recipient type.


## Example

The following Visual Basic for Applications (VBA) example creates a new **MailItem** object and adds Jon Grande as the recipient by using the default type ("To").

```vb
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande")
```

The following VBA example creates the same **MailItem** object as the preceding example, and then changes the type of the **Recipient** object from the default (To) to CC.

```vb
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande") 
 
myRecipient.Type = olCC
```

## Methods

|Name|
|:-----|
|[Delete](Outlook.Recipient.Delete.md)|
|[FreeBusy](Outlook.Recipient.FreeBusy.md)|
|[Resolve](Outlook.Recipient.Resolve.md)|

## Properties

|Name|
|:-----|
|[Address](Outlook.Recipient.Address.md)|
|[AddressEntry](Outlook.Recipient.AddressEntry.md)|
|[Application](Outlook.Recipient.Application.md)|
|[AutoResponse](Outlook.Recipient.AutoResponse.md)|
|[Class](Outlook.Recipient.Class.md)|
|[DisplayType](Outlook.Recipient.DisplayType.md)|
|[EntryID](Outlook.Recipient.EntryID.md)|
|[Index](Outlook.Recipient.Index.md)|
|[MeetingResponseStatus](Outlook.Recipient.MeetingResponseStatus.md)|
|[Name](Outlook.Recipient.Name.md)|
|[Parent](Outlook.Recipient.Parent.md)|
|[PropertyAccessor](Outlook.Recipient.PropertyAccessor.md)|
|[Resolved](Outlook.Recipient.Resolved.md)|
|[Sendable](Outlook.Recipient.Sendable.md)|
|[Session](Outlook.Recipient.Session.md)|
|[TrackingStatus](Outlook.Recipient.TrackingStatus.md)|
|[TrackingStatusTime](Outlook.Recipient.TrackingStatusTime.md)|
|[Type](Outlook.Recipient.Type.md)|

## See also

- [Recipient object members](overview/Outlook.md)
- [Obtain the email address of a recipient](../outlook/Concepts/Address-Book/obtain-the-e-mail-address-of-a-recipient.md)
- [Outlook Object Model reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
