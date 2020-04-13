---
title: Store.GetSearchFolders method (Outlook)
keywords: vbaol11.chm807
f1_keywords:
- vbaol11.chm807
ms.prod: outlook
api_name:
- Outlook.Store.GetSearchFolders
ms.assetid: aed6ba0b-5e20-adb9-6f62-d030a0de2e0b
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.GetSearchFolders method (Outlook)

Returns a **[Folders](Outlook.Folders.md)** collection object that represents the search folders defined for the **[Store](Outlook.Store.md)** object.


## Syntax

_expression_. `GetSearchFolders`

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Return value

A **Folders** collection object that represents all the search folders for the **Store** object.


## Remarks

 **GetSearchFolders** returns all the visible active search folders for the **Store**. It does not return uninitialized or aged out search folders.

 **GetSearchFolders** returns a **Folders** collection object with **[Folders.Count](Outlook.Folders.Count.md)** equal zero (0) if no search folders have been defined for the **Store**.

For a **Folders** collection object that represents a collection of search folders, **[Folders.Parent](Outlook.Folders.Parent.md)** returns the same object as **[Store.GetRootFolder](Outlook.Store.GetRootFolder.md)**. **[Folder.Folders](Outlook.Folder.Folders.md)** returns **Null** (**Nothing** in Visual Basic).


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) enumerates the search folders on all stores for the current session.


```vb
Sub EnumerateSearchFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oSearchFolders As Outlook.folders 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oSearchFolders = oStore.GetSearchFolders 
 
 For Each oFolder In oSearchFolders 
 
 Debug.Print (oFolder.FolderPath) 
 
 Next 
 
 Next 
 
End Sub
```


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]