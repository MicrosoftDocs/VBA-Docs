---
title: Folder.ShowItemCount property (Outlook)
keywords: vbaol11.chm2015
f1_keywords:
- vbaol11.chm2015
ms.prod: outlook
api_name:
- Outlook.Folder.ShowItemCount
ms.assetid: 3ce32c47-5f92-82ca-5ac3-b3d6f24e5f36
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.ShowItemCount property (Outlook)

Sets or returns a constant in the  **[OlShowItemCount](Outlook.OlShowItemCount.md)** enumeration that indicates whether to display the number of unread messages in the folder or the total number of items in the folder in the navigation pane. Read/write.


## Syntax

_expression_. `ShowItemCount`

_expression_ A variable that represents a [Folder](Outlook.Folder.md) object.


## Remarks

The  **ShowItemCount** property does not work with Public Folders.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays the unread count for the Inbox in the navigation pane.


```vb
Sub ShowTotalItemCount() 
 
 Dim nmsName As Outlook.NameSpace 
 
 Dim fldFolder As Outlook.Folder 
 
 
 
 Set nmsName = Application.GetNamespace("MAPI") 
 
 Set fldFolder = nmsName.GetDefaultFolder(olFolderInbox) 
 
 fldFolder.ShowItemCount = olShowUnreadItemCount 
 
End Sub
```


## See also


[Folder Object](Outlook.Folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]