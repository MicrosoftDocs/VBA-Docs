---
title: SharedWorkspaceFiles.Add method (Office)
keywords: vbaof11.chm267003
f1_keywords:
- vbaof11.chm267003
ms.prod: office
api_name:
- Office.SharedWorkspaceFiles.Add
ms.assetid: d6a8e86b-2075-be56-3e3f-75c3ffa6241c
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceFiles.Add method (Office)

Adds a file to the document library in a shared workspace. Returns a  **[SharedWorkspaceFile](Office.SharedWorkspaceFile.md)** object.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_. `Add`( `_FileName_`, `_ParentFolder_`, `_OverwriteIfFileAlreadyExists_`, `_KeepInSync_` )

 _expression_ Required. A variable that represents a '[SharedWorkspaceFiles](Office.SharedWorkspaceFiles.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The path and filename of the file to be added to the current shared workspace.|
| _ParentFolder_|Optional|**SharedWorkspaceFolder**|The subfolder in which to place the file, if not the main document library folder within the shared workspace. Add the file to the main document library folder by leaving this optional argument empty.|
| _OverwriteIfFileAlreadyExists_|Optional|**Boolean**|**True** to overwrite an existing file by the same name. Default is **False**.|
| _KeepInSync_|Optional|**Boolean**|**True** to keep the local copy of the document synchronized with the copy in the shared workspace. Default is **False**.|

## Example

The following example adds a new file to the files collection of the shared workspace.


```vb
    Dim swsfile As Office.SharedWorkspaceFile 
    Set swsfile = ActiveWorkbook.SharedWorkspace.Files.Add( _ 
        "C:\MyWorkbook.xls", _ 
        , True, True) 
    MsgBox "New file URL: " &amp; swsfile.URL, _ 
        vbInformation + vbOKOnly, _ 
        "New File in Shared Workspace Files" 
    Set swsfile = Nothing 

```


## See also


[SharedWorkspaceFiles Object](Office.SharedWorkspaceFiles.md)



[SharedWorkspaceFiles Object Members](./overview/Library-Reference/sharedworkspacefiles-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]