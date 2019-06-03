---
title: SharedWorkspaceTasks object (Office)
keywords: vbaof11.chm265000
f1_keywords:
- vbaof11.chm265000
ms.prod: office
api_name:
- Office.SharedWorkspaceTasks
ms.assetid: de26341f-44d1-131e-1dbe-e31f3f68e312
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceTasks object (Office)

A collection of the **[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)** objects in the current shared workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the **[Tasks](Office.SharedWorkspace.Tasks.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceTasks** collection.


```vb
    Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " & swsTasks.Count & _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## See also

- [SharedWorkspaceTasks object members](overview/Library-Reference/sharedworkspacetasks-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]