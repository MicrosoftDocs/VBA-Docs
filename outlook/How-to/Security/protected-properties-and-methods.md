---
title: Protected Properties and Methods
ms.prod: outlook
ms.assetid: 8522d350-a257-2924-2260-3cc02b6ebbca
ms.date: 06/08/2017
localization_priority: Normal
---


# Protected Properties and Methods

This topic lists the properties and methods in the Outlook object model that are protected by the Object Model Guard. If untrusted code performs a get on these properties or uses these methods, under default conditions for how Outlook is set up, it will invoke a security warning. The user will then have to verify and respond to the warning in order to proceed.

There are three security prompts that an untrusted application can possibly invoke, depending on the protected property or method that the application uses:

- The address book warning. This is the most common of the three security prompts. Unless marked otherwise, the properties and methods in the table below generate the address book warning.
    
- The execute action warning. Properties and methods superscripted by 1 in the table below denote that they generate the execute action warning.
    
- The send message warning. Properties and methods superscripted by 2 in the table below denote that they generate the send message warning.
    
For more information on security warnings, see  [Outlook Object Model Security Prompts](outlook-object-model-security-warnings.md).



| **Object**| **Protected Properties**| **Protected Methods**|
|:-----|:-----|:-----|
| [Account](../../../api/Outlook.Account.md)|CurrentUser, SmtpAddress|GetAddressEntryFromID, GetRecipientFromID|
| [Action](../../../api/Outlook.Action.md)||Execute1|
| [AddressEntries](../../../api/Outlook.AddressEntries.md)||Add, GetFirst, GetLast, GetNext, GetPrevious, Item|
| [AddressEntry](../../../api/Outlook.AddressEntry.md)|Address, ID, Manager, Members, Parent, PropertyAccessor|GetExchangeDistributionList, GetExchangeUser, Update|
| [AddressList](../../../api/Outlook.AddressList.md)|AddressEntries, ID, PropertyAccessor||
| [AddressLists](../../../api/Outlook.AddressLists.md)||Item|
| [AppointmentItem](../../../api/Outlook.AppointmentItem.md)|Body, OptionalAttendees, Organizer, PropertyAccessor, RequiredAttendees, Resources, RTFBody|Respond2, SaveAs, Send2|
| [Attachment](../../../api/Outlook.Attachment.md)|PropertyAccessor||
| [CalendarSharing](../../../api/Outlook.CalendarSharing.md)||SaveAsICal|
| [Columns](../../../api/Outlook.Columns.md)||Add|
| [ContactItem](../../../api/Outlook.ContactItem.md)|Body, Email1Address, Email1AddressType, Email1DisplayName, Email1EntryID, Email2Address, Email2AddressType, Email2DisplayName, Email2EntryID, Email3Address, Email3AddressType, Email3DisplayName, Email3EntryID, IMAddress, NetMeetingAlias, PropertyAccessor, ReferredBy, RTFBody|SaveAs|
| [DistListItem](../../../api/Outlook.DistListItem.md)|Body, PropertyAccessor, RTFBody|GetMember, SaveAs|
| [DocumentItem](../../../api/Outlook.DocumentItem.md)|Body, PropertyAccessor||
| [ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)|Address, Alias, ID, Parent, PrimarySmtpAddress, PropertyAccessor|GetExchangeDistributionList, GetExchangeUser, GetMemberOfList, GetExchangeDistributionListMembers, GetOwners, Update|
| [ExchangeUser](../../../api/Outlook.ExchangeUser.md)|Address, Alias, ID, Parent, PrimarySmtpAddress, PropertyAccessor|GetDirectReports, GetExchangeDistributionList, GetExchangeUser, GetExchangeUserManager, GetMemberOfList, Update|
| [Folder](../../../api/Outlook.Folder.md)|GetCalendarExporter, PropertyAccessor||
| [Inspector](../../../api/Outlook.Inspector.md)|HTMLEditor, WordEditor||
| [ItemProperties](../../../api/Outlook.ItemProperties.md)|Any protected property for an item||
| [JournalItem](../../../api/Outlook.JournalItem.md)|Body, ContactNames, PropertyAccessor|SaveAs|
| [MailItem](../../../api/Outlook.MailItem.md)|Bcc, Body, Cc, HTMLBody, PropertyAccessor, ReceivedByName, ReceivedOnBehalfOfName, Recipients, ReplyRecipientNames, RTFBody, Sender, SenderEmailAddress, SenderEmailType, SenderName, SentOnBehalfOfName, To|SaveAs, Send2|
| [MeetingItem](../../../api/Outlook.MeetingItem.md)|Body, PropertyAccessor, Recipients, RTFBody, SenderName|SaveAs|
| [MobileItem](../../../api/overview/Outlook.md)|Body, HTMLBody, PropertyAccessor, ReceivedByName, Recipients, ReplyRecipientNames, SenderEmailAddress, SenderEmailType, SenderName, SMILBody, To|SaveAs, Send2|
| [NameSpace](../../../api/Outlook.NameSpace.md)|CurrentUser, SelectNamesDialog|GetAddressEntryFromID, GetRecipientFromID|
| [NoteItem](../../../api/Outlook.NoteItem.md)|Body, PropertyAccessor||
| [PostItem](../../../api/Outlook.PostItem.md)|Body, HTMLBody, PropertyAccessor, RTFBody, SenderName|SaveAs|
| [Recipient](../../../api/Outlook.Recipient.md)|Any property|Any method|
| [Recipients](../../../api/Outlook.Recipients.md)|Any property|Any method|
| [RemoteItem](../../../api/Outlook.RemoteItem.md)|Body, PropertyAccessor||
| [ReportItem](../../../api/Outlook.ReportItem.md)|Body, PropertyAccessor||
| [SelectNamesDialog](../../../api/Outlook.SelectNamesDialog.md)|Recipients||
| [SharingItem](../../../api/Outlook.SharingItem.md)|Bcc, Body, Cc, HTMLBody, PropertyAccessor, ReceivedByName, ReceivedOnBehalfOfName, ReplyRecipientNames, RTFBody, SenderEmailAddress, SenderEmailType, SenderName, SendOnBehalfOfName, To|Allow, SaveAs, Send2|
| [StorageItem](../../../api/Outlook.StorageItem.md)|Body, PropertyAccessor||
| [Store](../../../api/Outlook.Store.md)|PropertyAccessor||
| [TaskItem](../../../api/Outlook.TaskItem.md)|Body, ContactNames, Contacts, Delegator, Owner, PropertyAccessor, RTFBody, StatusOnCompletionRecipients, StatusUpdateRecipients|SaveAs, Send2|
| [TaskRequestAcceptItem](../../../api/Outlook.TaskRequestAcceptItem.md)|Body, PropertyAccessor, RTFBody||
| [TaskRequestDeclineItem](../../../api/Outlook.TaskRequestDeclineItem.md)|Body, PropertyAccessor, RTFBody||
| [TaskRequestItem](../../../api/Outlook.TaskRequestItem.md)|Body, PropertyAccessor, RTFBody||
| [TaskRequestUpdateItem](../../../api/Outlook.TaskRequestUpdateItem.md)|Body, PropertyAccessor, RTFBody||
| [UserProperties](../../../api/Outlook.UserProperties.md)||Find|
| [UserProperty](../../../api/Outlook.UserProperty.md)|Formula||


> [!NOTE] 
> **[UserProperties.Find](../../../api/Outlook.UserProperties.Find.md)** is protected if the property being requested is one of the built-in properties that contains address information. If you ask for a custom property or a property like **Subject** that doesn't contain address information, a prompt will not be displayed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]