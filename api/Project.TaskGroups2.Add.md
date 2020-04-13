---
title: TaskGroups2.Add method (Project)
ms.prod: project-server
api_name:
- Project.TaskGroups2.Add
ms.assetid: 2f7a39a4-527f-1355-f3d0-4d5e674bf00c
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskGroups2.Add method (Project)

Adds a **[Group2](Project.Group2.md)** object to the **TaskGroups2** collection.


## Syntax

_expression_.**Add** (_Name_, _FieldName_)

_expression_ An expression that returns a 'TaskGroups2' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**| The name of a group definition.|
| _FieldName_|Required|**String**|The name of the first field to group by.|

## Return value

 **Group2**


## See also


[TaskGroups2 Collection Object](Project.taskgroups2(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]