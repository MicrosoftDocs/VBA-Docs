---
title: Stores object (Outlook)
keywords: vbaol11.chm3019
f1_keywords:
- vbaol11.chm3019
ms.prod: outlook
api_name:
- Outlook.Stores
ms.assetid: 8915a8e4-9c22-21d5-c492-051d393ce5f7
ms.date: 06/08/2017
localization_priority: Normal
---


# Stores object (Outlook)

A set of  **[Store](Outlook.Store.md)** objects representing all the stores available in the current profile.


## Remarks

You can use the  **Stores** and **Store** objects to enumerate all folders and search folders on all stores in the current session. For more information on storing Outlook items in folders and stores, see [Storing Outlook Items](../outlook/How-to/Items-Folders-and-Stores/storing-outlook-items.md).


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


## Events



|Name|
|:-----|
|[BeforeStoreRemove](Outlook.Stores.BeforeStoreRemove.md)|
|[StoreAdd](Outlook.Stores.StoreAdd.md)|

## Methods



|Name|
|:-----|
|[Item](Outlook.Stores.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Stores.Application.md)|
|[Class](Outlook.Stores.Class.md)|
|[Count](Outlook.Stores.Count.md)|
|[Parent](Outlook.Stores.Parent.md)|
|[Session](Outlook.Stores.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Stores Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]