---
title: MailItem.MarkForDownload property (Outlook)
keywords: vbaol11.chm1376
f1_keywords:
- vbaol11.chm1376
ms.prod: outlook
api_name:
- Outlook.MailItem.MarkForDownload
ms.assetid: 7ab16b80-90c6-ef60-b1ce-95fe87ab0d06
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.MarkForDownload property (Outlook)

Returns or sets an **[OlRemoteStatus](Outlook.OlRemoteStatus.md)** constant that determines the status of an item once it is received by a remote user. Read/write.


## Syntax

_expression_. `MarkForDownload`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


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


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]