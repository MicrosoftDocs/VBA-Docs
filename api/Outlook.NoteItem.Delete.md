---
title: NoteItem.Delete method (Outlook)
keywords: vbaol11.chm1494
f1_keywords:
- vbaol11.chm1494
ms.prod: outlook
api_name:
- Outlook.NoteItem.Delete
ms.assetid: 8af18879-a4e7-1ccf-bbc3-5140f506add1
ms.date: 06/08/2017
localization_priority: Normal
---


# NoteItem.Delete method (Outlook)

Removes the item from the folder that contains the item.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a [NoteItem](Outlook.NoteItem.md) object.


## Remarks

The  **Delete** method deletes a single item in a collection. To delete all items in the **[Items](Outlook.Folder.Items.md)** collection of a folder, you must delete each item starting with the last item in the folder. For example, in the items collection of a folder, `AllItems`, if there are  `n` number of items in the folder, start deleting the item at `AllItems.Item(n)`, decrementing the index each time until you delete  `AllItems.Item(1)`.

The  **Delete** method moves the item from the containing folder to the **Deleted Items** folder. If the containing folder is the **Deleted Items** folder, the **Delete** method removes the item permanently.


## See also


[NoteItemObject](Outlook.NoteItem.md)




[Delete All Items and Subfolders in the Deleted Items Folder](../outlook/How-to/Items-Folders-and-Stores/delete-all-items-and-subfolders-in-the-deleted-items-folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]