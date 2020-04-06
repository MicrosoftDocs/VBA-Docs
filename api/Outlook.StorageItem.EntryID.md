---
title: StorageItem.EntryID property (Outlook)
keywords: vbaol11.chm2143
f1_keywords:
- vbaol11.chm2143
ms.prod: outlook
api_name:
- Outlook.StorageItem.EntryID
ms.assetid: 5489c6df-8bd5-db6a-9d06-abe224813feb
ms.date: 06/08/2017
localization_priority: Normal
---


# StorageItem.EntryID property (Outlook)

Returns a  **String** representing the unique Entry ID of the object. Read-only.


## Syntax

_expression_. `EntryID`

_expression_ A variable that represents a [StorageItem](Outlook.StorageItem.md) object.


## Remarks

The EntryID is one of the three means to identify a  **[StorageItem](Outlook.StorageItem.md)** object using **[Folder.GetStorage](Outlook.Folder.GetStorage.md)**.

This property corresponds to the MAPI property  **PidTagEntryId**.

A MAPI store provider assigns a unique ID string when an item is created in its store. Therefore, the  **EntryID** property is not set for an Outlook item until it is saved or sent. The Entry ID changes when an item is moved into another store, for example, from your **Inbox** to a Microsoft Exchange Server public folder, or from one Personal Folders (.pst) file to another .pst file. Solutions should not depend on the **EntryID** property to be unique unless items will not be moved. The **EntryID** property returns a MAPI long-term Entry ID. For more information about long- and short-term EntryIDs, search https://msdn.microsoft.com for **PidTagEntryId**.


## See also


[StorageItem Object](Outlook.StorageItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]