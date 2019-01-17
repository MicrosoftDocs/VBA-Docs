---
title: SharedWorkspaceTasks object (Office)
keywords: vbaof11.chm265000
f1_keywords:
- vbaof11.chm265000
ms.prod: office
api_name:
- Office.SharedWorkspaceTasks
ms.assetid: de26341f-44d1-131e-1dbe-e31f3f68e312
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceTasks object (Office)

A collection of the  **[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)** objects in the current shared workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Tasks](Office.SharedWorkspace.Tasks.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceTasks** collection.


```vb
    Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " &amp; swsTasks.Count &amp; _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## Methods



|Name|
|:-----|
|[Add](Office.SharedWorkspaceTasks.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Office.SharedWorkspaceTasks.Application.md)|
|[Count](Office.SharedWorkspaceTasks.Count.md)|
|[Creator](Office.SharedWorkspaceTasks.Creator.md)|
|[Item](Office.SharedWorkspaceTasks.Item.md)|
|[ItemCountExceeded](Office.SharedWorkspaceTasks.ItemCountExceeded.md)|
|[Parent](Office.SharedWorkspaceTasks.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]