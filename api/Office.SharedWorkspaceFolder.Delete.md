---
title: SharedWorkspaceFolder.Delete Method (Office)
keywords: vbaof11.chm268006
f1_keywords:
- vbaof11.chm268006
ms.prod: office
api_name:
- Office.SharedWorkspaceFolder.Delete
ms.assetid: 188fff3c-4af9-4ebb-b846-9158cf7667e5
ms.date: 06/08/2017
---


# SharedWorkspaceFolder.Delete Method (Office)

Deletes the current shared workspace folder and all data within it.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. `Delete`( `_DeleteEvenIfFolderContainsFiles_` )

 _expression_ Required. A variable that represents a '[SharedWorkspaceFolder](Office.SharedWorkspaceFolder.md)' object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DeleteEvenIfFolderContainsFiles_|Optional|**Boolean**|**True** to delete the folder without warning even if the folder contains files. Default is **False**.The Delete method will fail if the user does not have permission to delete the current folder from the shared workspace.|

## See also


[SharedWorkspaceFolder Object](Office.SharedWorkspaceFolder.md)



[SharedWorkspaceFolder Object Members](sharedworkspacefolder-members-office.md)

