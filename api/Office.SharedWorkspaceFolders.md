---
title: SharedWorkspaceFolders object (Office)
keywords: vbaof11.chm269000
f1_keywords:
- vbaof11.chm269000
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders
ms.assetid: a9020edc-f199-6bab-75d1-c2bdc2a547d3
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFolders object (Office)

A collection of the **SharedWorkspaceFolder** objects in the current shared workspace.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the **Folders** property of the **SharedWorkspace** object to return a **SharedWorkspaceFolders** collection.


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

- [SharedWorkspaceFolders object members](overview/Library-Reference/sharedworkspacefolders-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]