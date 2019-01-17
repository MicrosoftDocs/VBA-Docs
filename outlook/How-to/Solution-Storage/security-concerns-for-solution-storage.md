---
title: Security Concerns for Solution Storage
ms.prod: outlook
ms.assetid: 8c237cd0-043a-d394-91a5-d85aab459091
ms.date: 06/08/2017
localization_priority: Normal
---


# Security Concerns for Solution Storage

This topic describes security considerations for storing private data in solution storage.

The Outlook object model intends  **[StorageItem](../../../api/Outlook.StorageItem.md)** objects to be created and accessed by only the solution or collaborating solutions that use them. Hence, it does not expose a **StorageItems** collection for all **StorageItem** objects in a folder. Custom properties created for the **StorageItem** are not exposed in the **Field Chooser** dialog box either.

The  **[Folder.GetTable](../../../api/Outlook.Folder.GetTable.md)** method supports a _TableContents_ parameter that returns a **[Table](../../../api/Outlook.Table.md)** containing only hidden items in a folder if you specify the parameter as **olHiddenItems**.

However, there exist technologies outside of the Outlook object model that allow modifying or deleting data stored as hidden items in MAPI folders. Solutions that are concerned with the privacy of their data should encrypt their private data at the property level with their own encryption algorithms.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]