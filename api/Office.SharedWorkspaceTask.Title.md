---
title: SharedWorkspaceTask.Title property (Office)
keywords: vbaof11.chm264001
f1_keywords:
- vbaof11.chm264001
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.Title
ms.assetid: 038d24fe-5afa-c61d-16e7-7a8c8fca2ccf
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceTask.Title property (Office)

Sets or gets the title of a **SharedWorkspaceTask** object. Read/write.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Title**

_expression_ A variable that represents a **[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)** object.


## Return value

String


## Remarks

The **Title** property is the single required property of a shared workspace task. Use the optional **Description** property to provide or return additional information about the task.


## Example

The following example displays a list of the titles of all tasks in the current shared workspace.


```vb
 Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTasks As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        strTasks = strTasks & swsTask.Title & vbCrLf 
    Next 
    MsgBox strTasks, vbInformation + vbOKOnly, _ 
        "Tasks in Shared Workspace" 
    Set swsTask = Nothing 
 

```


## See also

- [SharedWorkspaceTask object members](overview/Library-Reference/sharedworkspacetask-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]