---
title: TaskItem.DownloadState Property (Outlook)
keywords: vbaol11.chm1759
f1_keywords:
- vbaol11.chm1759
ms.prod: outlook
api_name:
- Outlook.TaskItem.DownloadState
ms.assetid: 65aa9f55-8b53-4c39-e560-c091d397e5ec
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.DownloadState Property (Outlook)

Returns a constant that belongs to the  **[OlDownloadState](Outlook.OlDownloadState.md)** enumeration indicating the download state of the item. Read-only.


## Syntax

_expression_. `DownloadState`

_expression_ A variable that represents a [TaskItem](./Outlook.TaskItem.md) object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example searches through the user's  **Inbox** for items that have not yet been fully downloaded. If any not yet fully downloaded items are found, a message is displayed to the user, and the item is marked for download.


```vb
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim objItems As Outlook.Items 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 Dim iCount As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 Set objItems = mpfInbox.Items 
 
 iCount = objItems.Count 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To iCount 
 
 Set obj = objItems.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox "This item has not been fully downloaded." 
 
 'Mark the item to be downloaded 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 obj.Save 
 
 End If 
 
 Next 
 
End Sub
```


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]