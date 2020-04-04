---
title: Store.IsOpen property (Outlook)
keywords: vbaol11.chm808
f1_keywords:
- vbaol11.chm808
ms.prod: outlook
api_name:
- Outlook.Store.IsOpen
ms.assetid: 05e93457-2d17-39ac-404c-c78c76d2ef72
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.IsOpen property (Outlook)

Returns a **Boolean** that indicates if the **[Store](Outlook.Store.md)** is open. Read-only.


## Syntax

_expression_.**IsOpen**

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Remarks

This property supports both Exchange and non-Exchange stores.

 **IsOpen** only indicates if the store is open. It does not indicate if the store is offline, or if it is an Exchange mailbox or an Exchange Public Folder and the store server is down.

Because opening a store can impose a performance overhead, and  **[Store.GetRootFolder](Outlook.Store.GetRootFolder.md)** and **[Store.GetSearchFolders](Outlook.Store.GetSearchFolders.md)** will open a store if it is not already open, you can use **IsOpen** before deciding to call **GetRootFolder** or **GetSearchFolders** to minimize performance overhead.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]