---
title: NameSpace object (Outlook)
keywords: vbaol11.chm3000
f1_keywords:
- vbaol11.chm3000
ms.prod: outlook
api_name:
- Outlook.NameSpace
ms.assetid: f0dcaa19-07f5-5d42-a3bf-2e42b7885644
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace object (Outlook)

Represents an abstract root object for any data source.


## Remarks

The object itself provides methods for logging in and out, accessing storage objects directly by ID, accessing certain special default folders directly, and accessing data sources owned by other users.

Use  **[GetNameSpace](Outlook.Application.GetNamespace.md)** ("MAPI") to return the Outlook **NameSpace** object from the **[Application](Outlook.Application.md)** object.

The only data source supported is MAPI, which allows access to all Outlook data stored in the user's mail stores.


## Events



|Name|
|:-----|
|[AutoDiscoverComplete](Outlook.NameSpace.AutoDiscoverComplete.md)|
|[OptionsPagesAdd](Outlook.NameSpace.OptionsPagesAdd.md)|

## Methods



|Name|
|:-----|
|[AddStore](Outlook.NameSpace.AddStore.md)|
|[AddStoreEx](Outlook.NameSpace.AddStoreEx.md)|
|[CompareEntryIDs](Outlook.NameSpace.CompareEntryIDs.md)|
|[CreateContactCard](Outlook.NameSpace.CreateContactCard.md)|
|[CreateRecipient](Outlook.NameSpace.CreateRecipient.md)|
|[CreateSharingItem](Outlook.NameSpace.CreateSharingItem.md)|
|[Dial](Outlook.NameSpace.Dial.md)|
|[GetAddressEntryFromID](Outlook.NameSpace.GetAddressEntryFromID.md)|
|[GetDefaultFolder](Outlook.NameSpace.GetDefaultFolder.md)|
|[GetFolderFromID](Outlook.NameSpace.GetFolderFromID.md)|
|[GetGlobalAddressList](Outlook.NameSpace.GetGlobalAddressList.md)|
|[GetItemFromID](Outlook.NameSpace.GetItemFromID.md)|
|[GetRecipientFromID](Outlook.NameSpace.GetRecipientFromID.md)|
|[GetSelectNamesDialog](Outlook.NameSpace.GetSelectNamesDialog.md)|
|[GetSharedDefaultFolder](Outlook.NameSpace.GetSharedDefaultFolder.md)|
|[GetStoreFromID](Outlook.NameSpace.GetStoreFromID.md)|
|[Logoff](Outlook.NameSpace.Logoff.md)|
|[Logon](Outlook.NameSpace.Logon.md)|
|[OpenSharedFolder](Outlook.NameSpace.OpenSharedFolder.md)|
|[OpenSharedItem](Outlook.NameSpace.OpenSharedItem.md)|
|[PickFolder](Outlook.NameSpace.PickFolder.md)|
|[RemoveStore](Outlook.NameSpace.RemoveStore.md)|
|[SendAndReceive](Outlook.NameSpace.SendAndReceive.md)|

## Properties



|Name|
|:-----|
|[Accounts](Outlook.NameSpace.Accounts.md)|
|[AddressLists](Outlook.NameSpace.AddressLists.md)|
|[Application](Outlook.NameSpace.Application.md)|
|[AutoDiscoverConnectionMode](Outlook.NameSpace.AutoDiscoverConnectionMode.md)|
|[AutoDiscoverXml](Outlook.NameSpace.AutoDiscoverXml.md)|
|[Categories](Outlook.NameSpace.Categories.md)|
|[Class](Outlook.NameSpace.Class.md)|
|[CurrentProfileName](Outlook.NameSpace.CurrentProfileName.md)|
|[CurrentUser](Outlook.NameSpace.CurrentUser.md)|
|[DefaultStore](Outlook.NameSpace.DefaultStore.md)|
|[ExchangeConnectionMode](Outlook.NameSpace.ExchangeConnectionMode.md)|
|[ExchangeMailboxServerName](Outlook.NameSpace.ExchangeMailboxServerName.md)|
|[ExchangeMailboxServerVersion](Outlook.NameSpace.ExchangeMailboxServerVersion.md)|
|[Folders](Outlook.NameSpace.Folders.md)|
|[Offline](Outlook.NameSpace.Offline.md)|
|[Parent](Outlook.NameSpace.Parent.md)|
|[Session](Outlook.NameSpace.Session.md)|
|[Stores](Outlook.NameSpace.Stores.md)|
|[SyncObjects](Outlook.NameSpace.SyncObjects.md)|
|[Type](Outlook.NameSpace.Type.md)|

## See also


[NameSpace Object Members](overview/Outlook.md)
[How to: Obtain and Log On to an Instance of Outlook](../outlook/How-to/Security/obtain-and-log-on-to-an-instance-of-outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
