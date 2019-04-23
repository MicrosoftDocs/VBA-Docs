---
title: Sync.ResolveConflict method (Office)
keywords: vbaof11.chm277010
f1_keywords:
- vbaof11.chm277010
ms.prod: office
api_name:
- Office.Sync.ResolveConflict
ms.assetid: d127ccab-644c-a2e3-68d1-57138ca200df
ms.date: 01/25/2019
localization_priority: Normal
---


# Sync.ResolveConflict method (Office)

Resolves conflicts between the local and the server copies of a shared document.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**ResolveConflict** (_SyncConflictResolution_)

_expression_ A variable that represents a **[Sync](Office.Sync.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SyncConflictResolution_|Required|**[MsoSyncConflictResolutionType](office.msosyncconflictresolutiontype.md)**|Specifies how conflicts should be resolved.|

## Remarks

Use the **ResolveConflict** method to resolve differences between the local copy of the active document and the server copy. Use the **msoSyncConflictMerge** option (not available for a Microsoft Excel Workbook) to merge the changes from each document into the other. Replace the server copy with local changes by using the **msoSyncConflictClientWins** option, or replace the local copy with the changed server copy by using the **msoSyncConflictServerWins** option.

The **msoSyncConflictMerge** option merges changes made to the server copy into the local copy, but does not actually resolve the conflict. To resolve the conflict with the merged changes winning, you must save the active document after merging changes, and then call the **ResolveConflict** method again with the **msoSyncConflictClientWins** option.

The **ResolveConflict** method can encounter a conflict condition if the client is unaware of recent changes to the server copy of the shared document. Call the **GetUpdate** method before calling **ResolveConflict** to refresh the status of the server copy and to detect a possible conflict.

The **ResolveConflict** method raises a run-time error if the local document has unsaved changes or if no conflict exists between the 2 copies of the document.

Not all document synchronization problems raise trappable run-time errors. After performing an operation by using the **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **[msoSyncStatusError](office.msosyncstatustype.md)**, check the **ErrorType** property for additional information about the type of error that has occurred.


## Example

The following example attempts to resolve a conflict by merging changes between the local and the server copies of the active document.


```vb
    Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    If objSync.Status = msoSyncStatusConflict Then 
        objSync.ResolveConflict msoSyncConflictMerge 
        ActiveDocument.Save 
        objSync.ResolveConflict msoSyncConflictClientWins 
        strStatus = "Conflict resolved by merging changes." 
        MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information" 
    End If 
    Set objSync = Nothing 

```


## See also

- [Sync object members](overview/Library-Reference/sync-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]