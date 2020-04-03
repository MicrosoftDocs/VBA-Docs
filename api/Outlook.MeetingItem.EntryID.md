---
title: MeetingItem.EntryID property (Outlook)
keywords: vbaol11.chm1411
f1_keywords:
- vbaol11.chm1411
ms.prod: outlook
api_name:
- Outlook.MeetingItem.EntryID
ms.assetid: 03d2684e-9608-f631-284d-ed63ce11c85a
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.EntryID property (Outlook)

Returns a **String** representing the unique Entry ID of the object. Read-only.


## Syntax

_expression_. `EntryID`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagEntryId**.

A MAPI store provider assigns a unique ID string when an item is created in its store. Therefore, the  **EntryID** property is not set for an Outlook item until it is saved or sent. The Entry ID changes when an item is moved into another store, for example, from your **Inbox** to a Microsoft Exchange Server public folder, or from one Personal Folders (.pst) file to another .pst file. Solutions should not depend on the **EntryID** property to be unique unless items will not be moved. The **EntryID** property returns a MAPI long-term Entry ID. For more information about long- and short-term EntryIDs, search https://msdn.microsoft.com for **PidTagEntryId**.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]