---
title: SharedWorkspaceFolders.Item property (Office)
keywords: vbaof11.chm269001
f1_keywords:
- vbaof11.chm269001
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders.Item
ms.assetid: 70916b0d-5cf7-b858-e215-d3cc948735fc
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFolders.Item property (Office)

Gets a **SharedWorkspaceFolder** object from the **Folders** collection of the shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Item**(_Index_)

_expression_ Required. A variable that represents a **[SharedWorkspaceFolders](Office.SharedWorkspaceFolders.md)** object. The specified **SharedWorkspaceFolders** collection.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Returns the **SharedWorkspaceFolder** at the position specified. The returned **SharedWorkspaceFolder** object does not correspond to the order in which the items are displayed in the **Shared Workspace** pane, and is not affected by re-sorting the display.|

## See also

- [SharedWorkspaceFolders object members](overview/Library-Reference/sharedworkspacefolders-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]