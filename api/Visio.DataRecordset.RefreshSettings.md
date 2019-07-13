---
title: DataRecordset.RefreshSettings property (Visio)
keywords: vis_sdr.chm16460345
f1_keywords:
- vis_sdr.chm16460345
ms.prod: visio
api_name:
- Visio.DataRecordset.RefreshSettings
ms.assetid: 7647676c-0291-8c57-10d6-ca55fcee2bf5
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.RefreshSettings property (Visio)

Gets and sets options that determine how the data recordset is refreshed. Read/write.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**RefreshSettings**

_expression_ An expression that returns a **[DataRecordset](Visio.DataRecordset.md)** object.


## Return value

Long


## Remarks

Constants for how a data recordset is refreshed are declared in the **VisRefreshSettings** enumeration in the Visio type library:



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRefreshNoReconciliationUI**|2|Disables the **Refresh Conflicts** task pane in the Visio user interface (UI) after a refresh operation.|
| **visRefreshOverwriteAll**|1|When data is refreshed, overwrites all user changes made in the shape data of shapes linked to data in this recordset since the previous refresh operation. See note.|

The default is for neither of the **VisRefreshSettings** flags to be turned on. (**RefreshSettings** = 0).

When **visRefreshNoReconciliationUI** is set, support for reconciling refresh conflicts in the Visio UI is disabled. As a developer, you should reconcile refresh conflicts programmatically by using the **[GetAllRefreshConflicts](Visio.DataRecordset.GetAllRefreshConflicts.md)**, **[GetMatchingRowsForRefreshConflict](Visio.DataRecordset.GetMatchingRowsForRefreshConflict.md)**, and **[RemoveRefreshConflict](Visio.DataRecordset.RemoveRefreshConflict.md)** methods.


> [!NOTE] 
> In some previous versions of Visio, shape data was called custom properties.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]