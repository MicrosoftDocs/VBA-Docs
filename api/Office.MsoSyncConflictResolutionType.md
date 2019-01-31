---
title: MsoSyncConflictResolutionType enumeration (Office)
ms.prod: office
api_name:
- Office.MsoSyncConflictResolutionType
ms.assetid: 2169c6ed-0460-3f6e-092a-d4a419be4525
ms.date: 01/31/2019
localization_priority: Normal
---


# MsoSyncConflictResolutionType enumeration (Office)

Specifies how conflicts should be resolved when synchronizing a shared document. Used with the **ResolveConflict** method of the **Sync** object.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
|**msoSyncConflictClientWins**|0|Replace the server copy with the local copy.|
|**msoSyncConflictMerge**|2|Merge changes made to the server copy into the local copy. To resolve the conflict with the merged changes winning, you must save the active document after merging changes, and then call the **ResolveConflict** method again with the **msoSyncConflictClientWins** option.|
|**msoSyncConflictServerWins**|1|Replace the local copy with the server copy.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]