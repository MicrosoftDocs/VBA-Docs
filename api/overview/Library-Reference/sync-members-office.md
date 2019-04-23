---
title: Sync members (Office)
description: The Sync property of the Document object in Word, the Workbook object in Excel, and the Presentation object in PowerPoint returns a Sync object.
ms.prod: office
ms.assetid: 748726bd-83de-425a-5af8-177c34e3a013
ms.date: 01/30/2019
localization_priority: Normal
---


# Sync members (Office)

The **Sync** property of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in Microsoft PowerPoint returns a **Sync** object.


## Methods

|Name|Description|
|:-----|:-----|
|[GetUpdate](../../Office.Sync.GetUpdate.md)|Compares the local version of the shared document to the version on the server.|
|[OpenVersion](../../Office.Sync.OpenVersion.md)|Opens a different version of the shared document alongside the currently open local version.|
|[PutUpdate](../../Office.Sync.PutUpdate.md)|Updates the server copy of the shared document with the local copy.|
|[ResolveConflict](../../Office.Sync.ResolveConflict.md)|Resolves conflicts between the local and the server copies of a shared document.|
|[Unsuspend](../../Office.Sync.Unsuspend.md)|Resumes synchronization between the local copy and the server copy of a shared document.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.Sync.Application.md)|Gets an **Application** object that represents the container application for the **Sync** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Creator](../../Office.Sync.Creator.md)|Gets a 32-bit integer that indicates the application in which the **Sync** object was created. Read-only.|
|[ErrorType](../../Office.Sync.ErrorType.md)|Gets a **MsoSyncErrorType** constant which indicates the type of the most recent document synchronization error. Read-only.|
|[LastSyncTime](../../Office.Sync.LastSyncTime.md)|Gets the date and time when the local copy of the active document was last synchronized with the server copy. Read-only.|
|[Parent](../../Office.Sync.Parent.md)|Gets the **Parent** object for the **Sync** object. Read-only.|
|[Status](../../Office.Sync.Status.md)|Gets the status of the synchronization of the local copy of the active document with the server copy. Read-only.|
|[WorkspaceLastChangedBy](../../Office.Sync.WorkspaceLastChangedBy.md)|Displays the display name of the user who last saved changes to the server copy of a shared document. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]