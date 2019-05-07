---
title: Task.OutlineHideSubTasks method (Project)
ms.prod: project-server
api_name:
- Project.Task.OutlineHideSubTasks
ms.assetid: 877e8248-3e3f-1816-0799-52fb5cda1d60
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.OutlineHideSubTasks method (Project)

Hides the subtasks of the selected task or tasks.


## Syntax

_expression_. `OutlineHideSubTasks`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Example

The following example collapses the entire outline of the first task.


```vb
Sub OutlineHideAllSubtasks() 
 ActiveProject.Tasks(1).OutlineHideSubtasks 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]