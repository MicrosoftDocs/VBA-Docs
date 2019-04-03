---
title: SyncObjects.AppFolders property (Outlook)
keywords: vbaol11.chm101
f1_keywords:
- vbaol11.chm101
ms.prod: outlook
api_name:
- Outlook.SyncObjects.AppFolders
ms.assetid: 711ebc16-12ac-9df3-31af-a883f438814f
ms.date: 06/08/2017
localization_priority: Normal
---


# SyncObjects.AppFolders property (Outlook)

This property returns the  **SyncObject** object for application folders. Read-only.


## Syntax

_expression_. `AppFolders`

_expression_ A variable that represents a [SyncObjects](Outlook.SyncObjects.md) object.


## Remarks

The  **SyncObject** is where folders are automatically added when the **InAppFolderSyncObject** property of the **Folder** object is set to **True**. The **SyncObject** allows users to synchronize Microsoft Outlook folders, address books, and folder home pages for offline use.


## Example

The following example sets the  **SyncObject** for the application folders and synchronizes the user's Inbox.


```vb
Public Sub SetAppfolders() 
 
 Dim nsp As Outlook.NameSpace 
 
 Dim objSycs As Outlook.SyncObjects 
 
 Dim objSyc As Outlook.SyncObject 
 
 Dim mpfInbox As Outlook.Folder 
 
 
 
 Set nsp = Application.GetNamespace("MAPI") 
 
 Set objSycs = nsp.SyncObjects 
 
 Set objSyc = objSycs.AppFolders 
 
 Set mpfInbox = nsp.GetDefaultFolder(olFolderInbox) 
 
 mpfInbox.InAppFolderSyncObject = True 
 
 objSyc.Start 
 
End Sub
```


## See also


[SyncObjects Object](Outlook.SyncObjects.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]