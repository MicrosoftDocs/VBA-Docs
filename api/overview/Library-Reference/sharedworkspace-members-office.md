---
title: SharedWorkspace members (Office)
ms.prod: office
ms.assetid: e4c2b518-d955-27e1-3e73-173d3c4f961d
ms.date: 01/30/2019
localization_priority: Normal
---


# SharedWorkspace members (Office)

The **SharedWorkspace** property of a **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **SharedWorkspace** object that allows the developer to add the active document to a SharePoint site and to manage other objects in the shared workspace site.


## Methods

|Name|Description|
|:-----|:-----|
|[CreateNew](../../Office.SharedWorkspace.CreateNew.md)|Creates a document workspace site on the server and adds the active document to the new shared workspace site.|
|[Delete](../../Office.SharedWorkspace.Delete.md)|Deletes the current shared workspace and all data within it.|
|[Disconnect](../../Office.SharedWorkspace.Disconnect.md)|Disconnects the local copy of the active document from the shared workspace site.|
|[Refresh](../../Office.SharedWorkspace.Refresh.md)|Refreshes the local cache of the **[SharedWorkspace](../../Office.SharedWorkspace.md)** object's files, folders, links, members, and tasks from the server.|
|[RemoveDocument](../../Office.SharedWorkspace.RemoveDocument.md)|Removes the active document from the shared workspace site.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.SharedWorkspace.Application.md)|Gets an **Application** object that represents the container application for the **SharedWorkspace** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Connected](../../Office.SharedWorkspace.Connected.md)|Gets a **Boolean** value that indicates whether or not the active document is currently saved in and connected to a shared workspace. Read-only.|
|[Creator](../../Office.SharedWorkspace.Creator.md)|Gets a 32-bit integer that indicates the application in which the **SharedWorkspace** object was created. Read-only.|
|[Files](../../Office.SharedWorkspace.Files.md)|Provides access to the **SharedWorkspaceFile** objects in the **SharedWorkspace**. Read-only.|
|[Folders](../../Office.SharedWorkspace.Folders.md)|Gets a **[SharedWorkspaceFolders](../../Office.SharedWorkspaceFolders.md)** collection that represents the list of subfolders in the document library associated with the current shared workspace. Read-only.|
|[LastRefreshed](../../Office.SharedWorkspace.LastRefreshed.md)|Gets the date and time when the **Refresh** method was most recently called. Read-only.|
|[Links](../../Office.SharedWorkspace.Links.md)|Gets a **[SharedWorkspaceLinks](../../Office.SharedWorkspaceLinks.md)** collection that represents the list of links saved in the current shared workspace. Read-only.|
|[Members](../../Office.SharedWorkspace.Members.md)|Gets a **[SharedWorkspaceMembers](../../Office.SharedWorkspaceMembers.md)** collection that represents the list of members in the current shared workspace. Read-only.|
|[Name](../../Office.SharedWorkspace.Name.md)|Gets or sets the display name of the shared workspace site. Read/write.|
|[Parent](../../Office.SharedWorkspace.Parent.md)|Gets the **Parent** object for the **SharedWorkspace** object. Read-only.|
|[SourceURL](../../Office.SharedWorkspace.SourceURL.md)|Designates the location of the public copy of a shared document to which changes should be published back after the document has been revised in a separate document workspace site. Read-only.|
|[Tasks](../../Office.SharedWorkspace.Tasks.md)|Gets a **[SharedWorkspaceTasks](../../Office.SharedWorkspaceTasks.md)** collection that represents the list of tasks in the current shared workspace. Read-only.|
|[URL](../../Office.SharedWorkspace.URL.md)|Gets the top-level Uniform Resource Locator (URL) of the shared workspace. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]