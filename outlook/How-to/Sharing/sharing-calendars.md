---
title: Sharing Calendars
ms.prod: outlook
ms.assetid: 03e0b693-5446-ca62-f868-69a583087966
ms.date: 06/08/2019
localization_priority: Normal
---


# Sharing Calendars

You can share calendar information in Microsoft Outlook by either sharing a calendar folder, if you have an Exchange Server account, or by exporting the contents of a calendar folder to an iCalendar calendar (.ics) file. Calendar information can also be shared online using a variety of protocols. 

For more information about sharing online calendars, see [Sharing Online Calendars, RSS Feeds, Microsoft SharePoint Foundation Folders, and Exchange Folders](sharing-online-calendars-rss-feeds-microsoft-sharepoint-foundation-folders-and-e.md).


## Sharing Calendar Folders

Sharing messages are used to either invite or request access to a calendar folder, or to respond to a sharing invitation or request by either allowing or denying access to a calendar folder. To construct a sharing invitation or sharing request, the **[CreateSharingItem](../../../api/Outlook.SharingItem.Recipients.md)** method of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object is used to create a **SharingItem** object. A **[Folder](../../../api/Outlook.Folder.md)** object reference to the desired calendar folder is used to establish the sharing context for the sharing message.

> [!NOTE] 
> You can only reference the **Calendar** default folder when creating a sharing request. A single **SharingItem** object can represent both a sharing invitation and a sharing request, if the **Calendar** default folder is used as the sharing context.

Sharing responses are automatically created and sent by calling the **[Allow](../../../api/Outlook.SharingItem.Allow.md)** or **[Deny](../../../api/Outlook.SharingItem.Deny.md)** methods of a **SharingItem** which represents a sharing request. Calling the **Allow** or **Deny** method allows or denies, respectively, access to the requested folder - the user requesting access need not receive the sharing response.


## Exporting Calendar Information

The **[CalendarSharing](../../../api/Outlook.CalendarSharing.md)** object is used to export information from the calendar folder to an iCalendar calendar file, and can also be used to create a **[MailItem](../../../api/Outlook.MailItem.md)** object that not only contains the iCalendar calendar file as an attachment, but also provides the calendar information as formatted HTML within the body of the mail message. The **CalendarSharing** object provides several properties that can be used to limit the scope and detail of calendar information included in the iCalendar calendar file and in the body of the MailItem.

The **[GetCalendarExporter](../../../api/Outlook.Folder.GetCalendarExporter.md)** method of the **Folder** object is used to obtain a **CalendarSharing** object reference for a specified calendar folder. From the **CalendarSharing** object, you can either use the **[SaveAsICal](../../../api/Outlook.CalendarSharing.SaveAsICal.md)** method to save an iCalendar calendar file, or you can use the **[ForwardAsICal](../../../api/Outlook.CalendarSharing.ForwardAsICal.md)** method to create a **MailItem** that contains both formatted calendar information and an iCalendar calendar file.

Once exported, an iCalendar calendar file can be opened by using the **[OpenSharedFolder](../../../api/Outlook.NameSpace.OpenSharedFolder.md)** method of the **NameSpace** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]