---
title: SharedWorkspaceFolder object (Office)
keywords: vbaof11.chm268000
f1_keywords:
- vbaof11.chm268000
ms.prod: office
api_name:
- Office.SharedWorkspaceFolder
ms.assetid: 297c4ed7-2232-5240-ca34-d374038c66a2
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceFolder object (Office)

Represents a folder in a shared document workspace.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **SharedWorkspaceFolder** object to manage subfolders within the main document library folder of a shared workspace.

 The **Count** property of the ** SharedWorkspaceFolders** collection does not include the workspace's main folder and returns 0 (zero) if no subfolders have been created.

The  **SharedWorkspaceFolder** object does not expose the **CreatedBy**, **CreatedDate**, **ModifiedBy**, and **ModifiedDate** properties available on the **SharedWorkspaceFile**, **SharedWorkspaceLink**, and **SharedWorkspaceTask** objects.

Use the  **Item** ( _index_ ) property of the **SharedWorkspaceFolders** collection to return a specific **SharedWorkspaceFolder** object.


## Example

Use the  **FolderName** property to return the name of the shared workspace folder. The following example returns the name of the first subfolder in the **SharedWorkspaceFolders** collection in the format "parentfoldername/foldername."


```vb
    Dim swsFolder As SharedWorkspaceFolder 
    Set swsFolder = ActiveWorkbook.SharedWorkspace.Folders(1) 
    MsgBox swsFolder.FolderName, vbInformation + vbOKOnly, "Folder Name" 
    Set swsFolder = Nothing 

```


## Methods



|Name|
|:-----|
|[Delete](Office.SharedWorkspaceFolder.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Office.SharedWorkspaceFolder.Application.md)|
|[Creator](Office.SharedWorkspaceFolder.Creator.md)|
|[FolderName](Office.SharedWorkspaceFolder.FolderName.md)|
|[Parent](Office.SharedWorkspaceFolder.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]