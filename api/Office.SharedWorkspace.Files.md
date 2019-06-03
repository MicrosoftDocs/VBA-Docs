---
title: SharedWorkspace.Files property (Office)
keywords: vbaof11.chm276004
f1_keywords:
- vbaof11.chm276004
ms.prod: office
api_name:
- Office.SharedWorkspace.Files
ms.assetid: e4a2f80e-5cb7-8ff2-3ab7-2b8c2d9d3cfb
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Files property (Office)

Provides access to the **SharedWorkspaceFile** objects in the **SharedWorkspace**. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Files**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Example

The following example lists the files saved in the current shared workspace.


```vb
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " & swsFiles.Count & _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]