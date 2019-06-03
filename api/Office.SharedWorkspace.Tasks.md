---
title: SharedWorkspace.Tasks property (Office)
keywords: vbaof11.chm276003
f1_keywords:
- vbaof11.chm276003
ms.prod: office
api_name:
- Office.SharedWorkspace.Tasks
ms.assetid: 9f7fa28d-f442-cbec-de7c-9109cc3e6f2e
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Tasks property (Office)

Gets a **[SharedWorkspaceTasks](Office.SharedWorkspaceTasks.md)** collection that represents the list of tasks in the current shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Tasks**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Example

The following example lists the tasks in the current shared workspace.


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

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]