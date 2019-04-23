---
title: DocumentLibraryVersion members (Office)
ms.prod: office
ms.assetid: 81015690-f681-67e5-4ff7-329a95f78f3d
ms.date: 01/30/2019
localization_priority: Normal
---


# DocumentLibraryVersion members (Office)

The **DocumentLibraryVersion** object represents a single saved version of a shared document which has versioning enabled and which is stored in a document library on the server. Each **DocumentLibraryVersion** object is a member of the active document's **DocumentLibraryVersions** collection.


## Methods

|Name|Description|
|:-----|:-----|
|[Delete](../../Office.DocumentLibraryVersion.Delete.md)|Removes a document library version from the **DocumentLibraryVersions** collection.|
|[Open](../../Office.DocumentLibraryVersion.Open.md)|Opens the specified version of the shared document from the **DocumentLibraryVersions** collection in read-only mode.|
|[Restore](../../Office.DocumentLibraryVersion.Restore.md)|Restores a previous saved version of a shared document from the **DocumentLibraryVersions** collection.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.DocumentLibraryVersion.Application.md)|Gets an **Application** object that represents the container application for the **DocumentLibraryVersion** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Comments](../../Office.DocumentLibraryVersion.Comments.md)|Gets any optional comments associated with the specified version of the shared document. Read-only.|
|[Creator](../../Office.DocumentLibraryVersion.Creator.md)|Gets a 32-bit integer that indicates the application in which the **DocumentLibraryVersion** object was created. Read-only.|
|[Index](../../Office.DocumentLibraryVersion.Index.md)|Gets a **Long** representing the index number for a **DocumentLibraryVersion** object in the collection. Read-only.|
|[Modified](../../Office.DocumentLibraryVersion.Modified.md)|Gets the date and time at which the specified version of the shared document was last saved to the server. Read-only.|
|[ModifiedBy](../../Office.DocumentLibraryVersion.ModifiedBy.md)|Gets the name of the user who last saved the specified version of the shared document to the server. Read-only.|
|[Parent](../../Office.DocumentLibraryVersion.Parent.md)|Gets the **Parent** object for the **DocumentLibraryVersion** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]