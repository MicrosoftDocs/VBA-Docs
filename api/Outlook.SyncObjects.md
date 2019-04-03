---
title: SyncObjects object (Outlook)
keywords: vbaol11.chm94
f1_keywords:
- vbaol11.chm94
ms.prod: outlook
api_name:
- Outlook.SyncObjects
ms.assetid: 88e59f63-d834-b174-bbda-0af0cf2d0520
ms.date: 06/08/2017
localization_priority: Normal
---


# SyncObjects object (Outlook)

Contains a set of  **[SyncObject](Outlook.SyncObject.md)** objects representing the **Send/Receive** groups for a user.


## Remarks

Use the  **[SyncObjects](Outlook.NameSpace.SyncObjects.md)** property to return the **SyncObjects** object from a **[NameSpace](Outlook.NameSpace.md)** object.

The  **SyncObjects** object is read-only. You cannot add an item to the collection. However, note that you can add one **Send/Receive** group using the **AppFolders** property which will create a **Send/Receive** group called Application Folders.


## Example

The following Microsoft Visual Basic for Applications (VBA) example retrieves the  **SyncObjects** object for the MAPI **NameSpace** object.


```vb
Set mySyncObjects = Application.GetNameSpace("MAPI").SyncObjects
```


## Methods



|Name|
|:-----|
|[Item](Outlook.SyncObjects.Item.md)|

## Properties



|Name|
|:-----|
|[AppFolders](Outlook.SyncObjects.AppFolders.md)|
|[Application](Outlook.SyncObjects.Application.md)|
|[Class](Outlook.SyncObjects.Class.md)|
|[Count](Outlook.SyncObjects.Count.md)|
|[Parent](Outlook.SyncObjects.Parent.md)|
|[Session](Outlook.SyncObjects.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]