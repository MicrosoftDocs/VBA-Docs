---
title: SyncObject object (Outlook)
keywords: vbaol11.chm2984
f1_keywords:
- vbaol11.chm2984
ms.prod: outlook
api_name:
- Outlook.SyncObject
ms.assetid: 099865b6-767f-8022-6839-875624f284f7
ms.date: 06/08/2017
localization_priority: Normal
---


# SyncObject object (Outlook)

Represents a  **Send\Receive** group for a user.


## Remarks

A  **Send\Receive** group lets users configure different synchronization scenarios, selecting which folders and which filters apply.

Use the  **[Item](Outlook.SyncObjects.Item.md)** method to retrieve the **SyncObject** object from a **[SyncObjects](Outlook.SyncObjects.md)** object. Because the **[Name](Outlook.SyncObject.Name.md)** property is the default property of the **SyncObject** object, you can identify the group by name.

The  **SyncObject** object is read-only; you cannot change its properties or create new ones. However, note that you can add one **Send/Receive** group using the **[SyncObjects.AppFolders](Outlook.SyncObjects.AppFolders.md)** property which will create a **Send/Receive** group called **Application Folders**.


## Example

The following example retrieves a  **SyncObject** object by name.


```vb
Set mySyncObject = mySyncObjects.Item("Daily")
```


## Events



|Name|
|:-----|
|[OnError](Outlook.SyncObject.OnError.md)|
|[Progress](Outlook.SyncObject.Progress.md)|
|[SyncEnd](Outlook.SyncObject.SyncEnd.md)|
|[SyncStart](Outlook.SyncObject.SyncStart.md)|

## Methods



|Name|
|:-----|
|[Start](Outlook.SyncObject.Start.md)|
|[Stop](Outlook.SyncObject.Stop.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.SyncObject.Application.md)|
|[Class](Outlook.SyncObject.Class.md)|
|[Name](Outlook.SyncObject.Name.md)|
|[Parent](Outlook.SyncObject.Parent.md)|
|[Session](Outlook.SyncObject.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]