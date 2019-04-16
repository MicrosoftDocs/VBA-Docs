---
title: TaskGroups2.Copy method (Project)
ms.prod: project-server
api_name:
- Project.TaskGroups2.Copy
ms.assetid: 7afc3518-e5bb-52be-0a45-edb436381250
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskGroups2.Copy method (Project)

Makes a copy of a group definition for the  **TaskGroups2** collection and returns a reference to the **[Group2](Project.Group2.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _NewName_)

 _expression_ An expression that returns a 'TaskGroups2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

## Return value

 **Group2**


## See also


[TaskGroups2 Collection Object](Project.taskgroups2(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]