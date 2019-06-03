---
title: SharedWorkspaceFolder.FolderName property (Office)
keywords: vbaof11.chm268001
f1_keywords:
- vbaof11.chm268001
ms.prod: office
api_name:
- Office.SharedWorkspaceFolder.FolderName
ms.assetid: 1a5df8fc-0e9a-3e4e-675d-dff3fd3e7f2a
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFolder.FolderName property (Office)

Gets the name of a subfolder within the main document library folder of a shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**FolderName**

_expression_ A variable that represents a **[SharedWorkspaceFolder](Office.SharedWorkspaceFolder.md)** object.


## Remarks

The **FolderName** property returns the subfolder name in the format parentfolder/subfolder. For example, if the shared workspace contains a folder named `"Supporting Documents"`, the **FolderName** property returns `"Shared Documents/Supporting Documents"`.


## Example

The following example displays the number of subfolders in the shared workspace and their names.


```vb
    Dim swsFolder As Office.SharedWorkspaceFolder 
    Dim strFolderInfo As String 
    strFolderInfo = "The shared workspace contains " & _ 
        ActiveWorkbook.SharedWorkspace.Folders.Count & " folder(s)." & vbCrLf 
    If ActiveWorkbook.SharedWorkspace.Folders.Count > 0 Then 
        For Each swsFolder In ActiveWorkbook.SharedWorkspace.Folders 
            strFolderInfo = strFolderInfo & swsFolder.FolderName & vbCrLf 
        Next 
    End If 
    MsgBox strFolderInfo, vbInformation + vbOKOnly, _ 
        "Folders in Shared Workspace" 
    Set swsFolder = Nothing 

```


## See also

- [SharedWorkspaceFolder object members](overview/Library-Reference/sharedworkspacefolder-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]