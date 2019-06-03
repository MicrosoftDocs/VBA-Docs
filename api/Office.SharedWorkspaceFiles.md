---
title: SharedWorkspaceFiles object (Office)
keywords: vbaof11.chm267000
f1_keywords:
- vbaof11.chm267000
ms.prod: office
api_name:
- Office.SharedWorkspaceFiles
ms.assetid: 5e2937f7-f794-dffb-a1ec-69ea9a9e3546
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFiles object (Office)

A collection of the **[SharedWorkspaceFile](Office.SharedWorkspaceFile.md)** objects in the current shared workspace.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the **[Files](Office.SharedWorkspace.Files.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceFiles** collection.


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

- [SharedWorkspaceFiles object members](overview/Library-Reference/sharedworkspacefiles-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]