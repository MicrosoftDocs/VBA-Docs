---
title: SharedWorkspaceTask.ModifiedBy property (Office)
keywords: vbaof11.chm264009
f1_keywords:
- vbaof11.chm264009
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.ModifiedBy
ms.assetid: e18d400b-0e53-a599-e789-d47c78abec49
ms.date: 06/08/2017
---


# SharedWorkspaceTask.ModifiedBy property (Office)

Gets the name of the user who last modified the object. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. `ModifiedBy`

 _expression_ A variable that represents a [SharedWorkspaceTask](Office.SharedWorkspaceTask.md) object.


## Return value

String


## Remarks

For shared workspace objects, the  **ModifiedBy** property returns the display name stored in the **Name** property of the **SharedWorkspaceMember** object. The **SharedWorkspaceMember** object does not have a **ModifiedBy** property.


## See also


[SharedWorkspaceTask Object](Office.SharedWorkspaceTask.md)



[SharedWorkspaceTask Object Members](./overview/Library-Reference/sharedworkspacetask-members-office.md)

