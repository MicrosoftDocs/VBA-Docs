---
title: TaskRequestAcceptItem.MarkForDownload property (Outlook)
keywords: vbaol11.chm1813
f1_keywords:
- vbaol11.chm1813
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.MarkForDownload
ms.assetid: f2dd7f89-5683-f8e1-b4b3-f49b29ec3534
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.MarkForDownload property (Outlook)

Returns or sets an **[OlRemoteStatus](Outlook.OlRemoteStatus.md)** constant that determines the status of an item once it is received by a remote user. Read/write.


## Syntax

_expression_. `MarkForDownload`

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Remarks

This property gives remote users with less-than-ideal data-transfer capabilities increased messaging flexibility.


## Example

The following example searches through the user's  **Inbox** for items that have not yet been fully downloaded. If any items are found that are not fully downloaded, a message is displayed and the item is marked for download.


```vb
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To mpfInbox.Items.Count 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox ("This item has not been fully downloaded.") 
 
 'Mark the item to be downloaded. 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 End If 
 
 Next 
 
End Sub
```


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]