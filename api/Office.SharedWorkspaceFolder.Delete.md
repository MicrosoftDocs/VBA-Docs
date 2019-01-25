---
title: SharedWorkspaceFolder.Delete method (Office)
keywords: vbaof11.chm268006
f1_keywords:
- vbaof11.chm268006
ms.prod: office
api_name:
- Office.SharedWorkspaceFolder.Delete
ms.assetid: 188fff3c-4af9-4ebb-b846-9158cf7667e5
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFolder.Delete method (Office)

Deletes the current shared workspace folder and all data within it.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Delete** (_DeleteEvenIfFolderContainsFiles_)

_expression_ Required. A variable that represents a **[SharedWorkspaceFolder](Office.SharedWorkspaceFolder.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DeleteEvenIfFolderContainsFiles_|Optional|**Boolean**|**True** to delete the folder without warning even if the folder contains files. Default is **False**. The **Delete** method will fail if the user does not have permission to delete the current folder from the shared workspace.|

## See also

- [SharedWorkspaceFolder object members](overview/Library-Reference/sharedworkspacefolder-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]