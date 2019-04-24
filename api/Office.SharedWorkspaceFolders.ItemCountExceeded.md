---
title: SharedWorkspaceFolders.ItemCountExceeded property (Office)
keywords: vbaof11.chm269005
f1_keywords:
- vbaof11.chm269005
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders.ItemCountExceeded
ms.assetid: cc8f3b36-e9cc-ad08-c94d-85c2b909ee97
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFolders.ItemCountExceeded property (Office)

Gets a **Boolean** value that indicates whether the number of **SharedWorkspaceFolders** items in the collection has exceeded the 99 that can be displayed in the **Shared Workspace** task pane. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**ItemCountExceeded**

_expression_ A variable that represents a **[SharedWorkspaceFolders](Office.SharedWorkspaceFolders.md)** object.


## Return value

Boolean


## Remarks

The **Shared Workspace** task pane can only display 99 shared workspace files and folders, links, members, or tasks. If more than 99 items are added to any of these collections, the corresponding tab of the **Shared Workspace** task pane will stop displaying the list of items, and displays a link to the shared workspace site webpage instead; the collection is no longer populated locally, and its **Count** property returns 0 (zero).

Furthermore, after the **ItemCountExceeded** property returns **True** for one of the collections listed earlier, the developer can no longer remedy the situation programmatically by deleting items from the collection to reduce the count below 99 because the collection is no longer populated.

The **ItemCountExceeded** property of the **SharedWorkspaceFolders** collection returns **True** when the combined count of files and folders exceeds 99 because both lists are combined and displayed together on the **Documents** tab of the **Shared Workspace** task pane.


## See also

- [SharedWorkspaceFolders object members](overview/Library-Reference/sharedworkspacefolders-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]