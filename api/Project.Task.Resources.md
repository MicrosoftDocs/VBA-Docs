---
title: Task.Resources property (Project)
ms.prod: project-server
api_name:
- Project.Task.Resources
ms.assetid: 72f4535f-39f1-81eb-7400-47fbca9cccd4
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Resources property (Project)

Gets a  **[Resources](Project.Resource.md)** collection that contains the resources assigned to the task. Read-only **Resources**.


## Syntax

_expression_. `Resources`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Example

The following example displays the name of each resource assigned to the selected task.


```vb
Sub ResourceNames() 
 
 Dim R As Resource 
 
 For Each R In ActiveCell.Task.Resources 
 MsgBox R.Name 
 Next R 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]