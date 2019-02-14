---
title: MsoSyncStatusType enumeration (Office)
ms.prod: office
api_name:
- Office.MsoSyncStatusType
ms.assetid: 52dab603-eb05-709a-99d5-908f2713b953
ms.date: 01/31/2019
localization_priority: Normal
---


# MsoSyncStatusType enumeration (Office)

Specifies the status of the synchronization of the local copy of the active document with the server copy. Used with the **Status** property of the **Sync** object.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
|**msoSyncStatusConflict**|4|Both the local and the server copies have changes.|
|**msoSyncStatusError**|6|An error occurred. Use the **ErrorType** property of the **Sync** object to determine the exact error.|
|**msoSyncStatusLatest**|1|Documents are already in sync.|
|**msoSyncStatusLocalChanges**|3|Only local copy has changes.|
|**msoSyncStatusNewerAvailable**|2|Only server copy has changes.|
|**msoSyncStatusNoSharedWorkspace**|0|No shared workspace.|
|**msoSyncStatusNotRoaming**|0|No synchronization is needed.|
|**msoSyncStatusSuspended**|5|Synchronization has been suspended.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]