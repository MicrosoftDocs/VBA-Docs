---
title: SharedWorkspace.Folders property (Office)
keywords: vbaof11.chm276005
f1_keywords:
- vbaof11.chm276005
ms.prod: office
api_name:
- Office.SharedWorkspace.Folders
ms.assetid: aaba6357-fff5-f3d2-e7d7-6453183864e3
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Folders property (Office)

Gets a **[SharedWorkspaceFolders](Office.SharedWorkspaceFolders.md)** collection that represents the list of subfolders in the document library associated with the current shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Folders**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Remarks

The **SharedWorkspaceFolders** collection does not include the root document library folder itself, which by default is named `"Shared Documents"`.


## Example

The following example lists the subfolders in the current shared workspace.


```vb
    Dim swsFolders As Office.SharedWorkspaceFolders 
    Set swsFolders = ActiveWorkbook.SharedWorkspace.Folders 
    MsgBox "There are " & swsFolders.Count & _ 
        " folder(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFolders = Nothing 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]