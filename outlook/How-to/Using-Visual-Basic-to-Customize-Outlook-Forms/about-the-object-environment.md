---
title: About the Object Environment
keywords: olfm10.chm3077351
f1_keywords:
- olfm10.chm3077351
ms.prod: outlook
ms.assetid: 03aa62d6-23be-8a2e-73e2-b1ff6307545d
ms.date: 06/08/2019
localization_priority: Normal
---


# About the Object Environment

There are two ways to write code for Microsoft Outlook:


- From outside the application, such as by using Microsoft Visual Basic or Microsoft Visual Basic for Applications in Microsoft Excel or another application.
    
- From inside the application, such as by using Visual Basic for Applications or by using VBScript with an Outlook form.
    

## Major components of the Outlook object model

The following table shows the major objects in the Outlook object model.


| Component | Description | Example |
| ------- | ------------------------ | --------- |
| **[Application](../../../api/Outlook.Application.md)**|The top of the object hierarchy that represents the entire application. Enables you to reference other objects in the application and create items and objects. | This code creates an appointment in Outlook: `Application.CreateItem(1).Display`|
| **[NameSpace](../../../api/Outlook.NameSpace.md)**|Represents the MAPI message store where all the Outlook items are stored. Provides methods for logging on and off Outlook and for referencing the default folders such as Mailbox, Inbox, Contacts, and others. | This code references the active user in Outlook: `Application.GetNameSpace("MAPI").CurrentUser`|
| **[Account](../../../api/Outlook.Account.md)**|Represents an account defined for the current profile.| |
| **[Store](../../../api/Outlook.Store.md)**|Represents a file on the local computer or a network drive that stores email messages and other items for an account in the current profile. | |
| **[Folders](../../../api/Outlook.Folders.md)**| There are two folder objects, the **Folders** collection object that enables you to work with collections of folders and the **[Folder](../../../api/Outlook.Folder.md)** object that enables you to work with a single folder. | This code shows the collection of folders named Personal Folders in Outlook: `Application.GetNameSpace("MAPI").Folders("Personal Folders")`|
| **[Table](../../../api/Outlook.Table.md)**|Represents a set of item data from a **Folder** or **[Search](../../../api/Outlook.Search.md)** object, with items as rows of the table and properties as columns of the table.| |
| **[Rule](../../../api/Outlook.Rule.md)**|Represents an Outlook rule.| |
| **[View](../../../api/Outlook.View.md)**|Represents a customizable view used to sort, group, and view data.| |
| **[Explorer](../../../api/Outlook.Explorer.md)**|Represents the Outlook window. Enables you to show, return, and close the active window. | This code shows the active Outlook window in Outlook: `Application.ActiveExplorer.Display`|
| **[NavigationPane](../../../api/Outlook.NavigationPane.md)**|Represents the Navigation Pane displayed by the active **Explorer** object.| |
| **[Items](../../../api/Outlook.Items.md)** collection | Enables you to work with items within a folder and the item objects that represents the standard item types in Outlook, such as **[MailItem](../../../api/Outlook.MailItem.md)** that represents a mail message. In VBScript, the active item is assumed, so you do not need to enter the object model to reference it. | This code sets the Subject field of the active message in VBScript: `Item.Subject = "New Subject"`|
| **[Inspector](../../../api/Outlook.Inspector.md)**|References forms. Use to show forms and pages. | This code shows the **Options** page of a form in Outlook: `Application.ActiveInspector.SetCurrentFormPage("Options")`|
| **[FormRegion](../../../api/Outlook.FormRegion.md)**|Represents a form region in an Outlook form.| |
| **[Attachment](../../../api/Outlook.Attachment.md)**|Represents a document or link to a document contained in an Outlook item.| |
| **[PropertyAccessor](../../../api/Outlook.PropertyAccessor.md)**|Provides the ability to create, get, set, and delete properties on objects.| |
| **[ItemProperty](../../../api/Outlook.ItemProperty.md)**|Represents information about a given item property for an Outlook item object.| |
| **[UserProperty](../../../api/Outlook.UserProperty.md)**|Represents a custom property of an Outlook item.| |
| **[AddressEntry](../../../api/Outlook.AddressEntry.md)**|Each **AddressEntry** object in the **[AddressEntries](../../../api/Outlook.AddressEntries.md)** collection holds information that represents a person or process to which the messaging system can deliver messages.| |
| **[AddressList](../../../api/Outlook.AddressList.md)**|The **AddressList** object is an address book that contains a set of **AddressEntry** objects. The entire hierarchy is available through the parent **[AddressLists](../../../api/Outlook.AddressLists.md)** collection.| |
| **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)**|Provides detailed information about an **AddressEntry** that represents a Microsoft Exchange Server mailbox user.| |
| **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)**|Provides detailed information about an **AddressEntry** that represents an Exchange distribution list.| |
| **[Recipient](../../../api/Outlook.Recipient.md)**|Represents a user or resource in Outlook, generally a mail message addressee.| |
| **[Exception](../../../api/Outlook.Exception.md)**|The **Exception** object holds information about one instance of an **[AppointmentItem](../../../api/Outlook.AppointmentItem.md)** object which is an exception to a recurring series. Unlike most of the other Outlook objects, the **Exception** object is a read-only object.| |
|Control|There are the Microsoft Forms 2.0 controls that exist in the control toolbox by default, and the Outlook controls that are installed on your computer by default and that you will add to the control toolbox before using them for the first time in a form.| |

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]