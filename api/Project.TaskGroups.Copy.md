---
title: TaskGroups.Copy method (Project)
ms.prod: project-server
api_name:
- Project.TaskGroups.Copy
ms.assetid: e69fe06d-3855-a8ac-32fe-752ff280fe85
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskGroups.Copy method (Project)

Makes a copy of a group definition for the **TaskGroups** collection and returns a reference to the **[Group](Project.Group.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _NewName_)

_expression_ A variable that represents a 'TaskGroups' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the group to copy.|
| _NewName_|Required|**String**|The name of the new group.|

## Return value

 **Group**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]