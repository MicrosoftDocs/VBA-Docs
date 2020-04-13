---
title: Folder.StoreID property (Outlook)
keywords: vbaol11.chm1992
f1_keywords:
- vbaol11.chm1992
ms.prod: outlook
api_name:
- Outlook.Folder.StoreID
ms.assetid: 8b2657b7-0c69-d8ad-147b-482303ebd10f
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.StoreID property (Outlook)

Returns a **String** indicating the store ID for the folder. Read-only.


## Syntax

_expression_. `StoreID`

_expression_ A variable that represents a [Folder](Outlook.Folder.md) object.


## Example

This Visual Basic for Applications (VBA) example obtains the  **[EntryID](Outlook.Folder.EntryID.md)** and **StoreID** for the default Tasks folder and then calls the **[NameSpace.GetFolderFromID](Outlook.NameSpace.GetFolderFromID.md)** method using these values to obtain the same folder. The folder is then displayed.


```vb
Sub GetWithID() 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myEntryID As String 
 
 Dim myStoreID As String 
 
 Dim myNewFolder As Outlook.Folder 
 
 
 
 Set myFolder = Application.Session.GetDefaultFolder(olFolderTasks) 
 
 myEntryID = myFolder.EntryID 
 
 myStoreID = myFolder.StoreID 
 
 Set myNewFolder = Application.Session.GetFolderFromID(myEntryID, myStoreID) 
 
 myNewFolder.Display 
 
End Sub
```


## See also


[Folder Object](Outlook.Folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]