---
title: Task.UnlinkPredecessors method (Project)
ms.service: project-server
api_name:
- Project.Task.UnlinkPredecessors
ms.assetid: 2ac8703e-d282-d16a-e4b4-44dcd847cc6a
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Task.UnlinkPredecessors method (Project)

Removes one or more predecessors from the task.


## Syntax

_expression_.**UnlinkPredecessors**( `_Tasks_` )

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Tasks_|Required|**Object**|Can be a **Task** or **Tasks** object, which specifies one or more tasks that are removed as predecessors.|

## Return value

 **Nothing**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]