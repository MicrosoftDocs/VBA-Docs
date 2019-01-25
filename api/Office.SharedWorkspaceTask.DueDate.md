---
title: SharedWorkspaceTask.DueDate property (Office)
keywords: vbaof11.chm264006
f1_keywords:
- vbaof11.chm264006
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.DueDate
ms.assetid: 86ef146e-7528-9dfb-646f-8412abade012
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceTask.DueDate property (Office)

Gets or sets the optional due date and time of a **SharedWorkspaceTask** object. Read/write.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**DueDate** ()

_expression_ An expression that returns a **[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)** object.


## Example

The following example sets the **DueDate** of all tasks in a shared workspace to 12:00 noon on December 31, 2005, and uploads these changes to the server by using the **Save** method.


```vb
Dim swsTask As Office.SharedWorkspaceTask 
    Const dtmNewDueDate As Date = #12/31/2005 12:00:00 PM# 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        swsTask.DueDate = dtmNewDueDate 
        swsTask.Save 
    Next 
    Set swsTask = Nothing
```


## See also

- [SharedWorkspaceTask object members](overview/Library-Reference/sharedworkspacetask-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]