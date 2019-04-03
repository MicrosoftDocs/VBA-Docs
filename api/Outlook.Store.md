---
title: Store object (Outlook)
keywords: vbaol11.chm3155
f1_keywords:
- vbaol11.chm3155
ms.prod: outlook
api_name:
- Outlook.Store
ms.assetid: 1eb22fe9-8849-7476-5388-2515b48591b9
ms.date: 06/08/2017
localization_priority: Normal
---


# Store object (Outlook)

Represents a file on the local computer or a network drive that stores email messages and other items for an account in the current profile.


## Remarks

A profile defines one or more email accounts, and each email account is associated with a server of a specific type. For an Exchange server, a store can be on the server, in an Exchange Public folder, or in a local Personal Folders File (.pst) or Offline Folder File (.ost). For a POP3, IMAP, or HTTP email server, a store is a .pst file.

You can use the  **[Stores](Outlook.Stores.md)** and **Store** objects to enumerate all folders and search folders on all stores in the current session. Since getting the root folder or search folders in a store requires the store to be open and opening a store imposes an overhead on performance, you can check the **[Store.IsOpen](Outlook.Store.IsOpen.md)** property before you decide to pursue the operation.

If you use an Exchange server, you can access other explicit built-in  **Store** properties for store characteristics such as **[ExchangeStoreType](Outlook.Store.ExchangeStoreType.md)**, **[IsCachedExchange](Outlook.Store.IsCachedExchange.md)**, and **[IsDataFileStore](Outlook.Store.IsDataFileStore.md)**. Use the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object returned by **[Store.PropertyAccessor](Outlook.Store.PropertyAccessor.md)** to access other store properties that are not exposed in the Outlook object model.

For more information on storing Outlook items in folders and stores, see [Storing Outlook Items](../outlook/How-to/Items-Folders-and-Stores/storing-outlook-items.md).


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) enumerates all folders on all stores for a session:


```vb
Sub EnumerateFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oRoot As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oRoot = oStore.GetRootFolder 
 
 Debug.Print (oRoot.FolderPath) 
 
 EnumerateFolders oRoot 
 
 Next 
 
End Sub 
 
 
 
Private Sub EnumerateFolders(ByVal oFolder As Outlook.Folder) 
 
 Dim folders As Outlook.folders 
 
 Dim Folder As Outlook.Folder 
 
 Dim foldercount As Integer 
 
 
 
 On Error Resume Next 
 
 Set folders = oFolder.folders 
 
 foldercount = folders.Count 
 
 'Check if there are any folders below oFolder 
 
 If foldercount Then 
 
 For Each Folder In folders 
 
 Debug.Print (Folder.FolderPath) 
 
 EnumerateFolders Folder 
 
 Next 
 
 End If 
 
End Sub
```


## Methods



|Name|
|:-----|
|[GetDefaultFolder](Outlook.Store.GetDefaultFolder.md)|
|[GetRootFolder](Outlook.Store.GetRootFolder.md)|
|[GetRules](Outlook.Store.GetRules.md)|
|[GetSearchFolders](Outlook.Store.GetSearchFolders.md)|
|[GetSpecialFolder](Outlook.Store.GetSpecialFolder.md)|
|[RefreshQuotaDisplay](Outlook.Store.RefreshQuotaDisplay.md)|
|[CreateUnifiedGroup](Outlook.store.createunifiedgroup.md)|
|[DeleteUnifiedGroup](Outlook.store.deleteunifiedgroup.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Store.Application.md)|
|[Categories](Outlook.Store.Categories.md)|
|[Class](Outlook.Store.Class.md)|
|[DisplayName](Outlook.Store.DisplayName.md)|
|[ExchangeStoreType](Outlook.Store.ExchangeStoreType.md)|
|[FilePath](Outlook.Store.FilePath.md)|
|[IsCachedExchange](Outlook.Store.IsCachedExchange.md)|
|[IsConversationEnabled](Outlook.Store.IsConversationEnabled.md)|
|[IsDataFileStore](Outlook.Store.IsDataFileStore.md)|
|[IsInstantSearchEnabled](Outlook.Store.IsInstantSearchEnabled.md)|
|[IsOpen](Outlook.Store.IsOpen.md)|
|[Parent](Outlook.Store.Parent.md)|
|[PropertyAccessor](Outlook.Store.PropertyAccessor.md)|
|[Session](Outlook.Store.Session.md)|
|[StoreID](Outlook.Store.StoreID.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Store Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
