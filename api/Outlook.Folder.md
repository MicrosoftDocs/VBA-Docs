---
title: Folder object (Outlook)
keywords: vbaol11.chm3020
f1_keywords:
- vbaol11.chm3020
api_name:
- Outlook.Folder
ms.assetid: 3cf6cda8-6d70-666e-2643-9d9c5b9cacfc
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Folder object (Outlook)

Represents an Outlook folder.


## Remarks

A **Folder** object can contain other **Folder** objects, as well as Outlook items. Use the **Folders** property of a **[NameSpace](Outlook.NameSpace.md)** object or another **Folder** object to return the set of folders in a **NameSpace** or under a folder. You can navigate nested folders by starting from a top-level folder, say the Inbox, and using a combination of the **[Folder.Folders](Outlook.Folder.Folders.md)** property, which returns the set of folders underneath a **Folder** object in the hierarchy, and the **[Folders.Item](Outlook.Folders.Item.md)** method, which returns a folder within the **[Folders](Outlook.Folders.md)** collection.

There is a set of folders within an Outlook data store that supports the default functionality of Outlook. Use **[NameSpace.GetDefaultFolder](Outlook.NameSpace.GetDefaultFolder.md)**, specifying an _index_ that is one of the constants in the **[OlDefaultFolders](Outlook.OlDefaultFolders.md)** enumeration to return one of the default Outlook folders in the Outlook **NameSpace** object.

 While generally it is a good practice to place items that serve the same functionality in the same folder, a folder can contain items of different types. For example, by default, the Calendar folder can contain **[AppointmentItem](Outlook.AppointmentItem.md)** and **[MeetingItem](Outlook.MeetingItem.md)** objects, and the Contacts folder can contain **[ContactItem](Outlook.ContactItem.md)** and **[DistListItem](Outlook.DistListItem.md)** objects. In general, when enumerating items in a folder, don't assume the type of an item in the folder; check the message class of the item before accessing properties that are applicable to the item.

 Use the **[Folders.Add](Outlook.Folders.Add.md)** method to add a folder to the **Folders** object. The **Add** method has an optional argument that can be used to specify the type of items that can be stored in that folder. By default, folders created inside another folder inherit the type of the parent folder.

 Note that when items of a specific type are saved, they are saved directly into their corresponding default folder. For example, when the **[MeetingItem.GetAssociatedAppointment](Outlook.MeetingItem.GetAssociatedAppointment.md)** method is applied to a **MeetingItem** in the Inbox folder, the appointment that is returned will be saved to the default Calendar folder.


## Events



|Name|
|:-----|
|[BeforeFolderMove](Outlook.Folder.BeforeFolderMove.md)|
|[BeforeItemMove](Outlook.Folder.BeforeItemMove.md)|

## Methods



|Name|
|:-----|
|[AddToPFFavorites](Outlook.Folder.AddToPFFavorites.md)|
|[CopyTo](Outlook.Folder.CopyTo.md)|
|[Delete](Outlook.Folder.Delete.md)|
|[Display](Outlook.Folder.Display.md)|
|[GetCalendarExporter](Outlook.Folder.GetCalendarExporter.md)|
|[GetCustomIcon](Outlook.Folder.GetCustomIcon.md)|
|[GetExplorer](Outlook.Folder.GetExplorer.md)|
|[GetOwner](Outlook.Folder.GetOwner.md)|
|[GetStorage](Outlook.Folder.GetStorage.md)|
|[GetTable](Outlook.Folder.GetTable.md)|
|[MoveTo](Outlook.Folder.MoveTo.md)|
|[SetCustomIcon](Outlook.Folder.SetCustomIcon.md)|

## Properties



|Name|
|:-----|
|[AddressBookName](Outlook.Folder.AddressBookName.md)|
|[Application](Outlook.Folder.Application.md)|
|[Class](Outlook.Folder.Class.md)|
|[CurrentView](Outlook.Folder.CurrentView.md)|
|[CustomViewsOnly](Outlook.Folder.CustomViewsOnly.md)|
|[DefaultItemType](Outlook.Folder.DefaultItemType.md)|
|[DefaultMessageClass](Outlook.Folder.DefaultMessageClass.md)|
|[Description](Outlook.Folder.Description.md)|
|[EntryID](Outlook.Folder.EntryID.md)|
|[FolderPath](Outlook.Folder.FolderPath.md)|
|[Folders](Outlook.Folder.Folders.md)|
|[InAppFolderSyncObject](Outlook.Folder.InAppFolderSyncObject.md)|
|[IsSharePointFolder](Outlook.Folder.IsSharePointFolder.md)|
|[Items](Outlook.Folder.Items.md)|
|[Name](Outlook.Folder.Name.md)|
|[Parent](Outlook.Folder.Parent.md)|
|[PropertyAccessor](Outlook.Folder.PropertyAccessor.md)|
|[Session](Outlook.Folder.Session.md)|
|[ShowAsOutlookAB](Outlook.Folder.ShowAsOutlookAB.md)|
|[ShowItemCount](Outlook.Folder.ShowItemCount.md)|
|[Store](Outlook.Folder.Store.md)|
|[StoreID](Outlook.Folder.StoreID.md)|
|[UnReadItemCount](Outlook.Folder.UnReadItemCount.md)|
|[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)|
|[Views](Outlook.Folder.Views.md)|
|[WebViewOn](Outlook.Folder.WebViewOn.md)|
|[WebViewURL](Outlook.Folder.WebViewURL.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Folder Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
