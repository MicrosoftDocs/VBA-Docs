---
title: Task.Flag8 property (Project)
ms.prod: project-server
api_name:
- Project.Task.Flag8
ms.assetid: 50e174bb-93fd-074b-ded7-a49ef55d87b4
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Flag8 property (Project)

Gets or sets the value of a task flag custom field. Read/write  **Variant**.


## Syntax

_expression_. `Flag8`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Example

The following example deletes all the tasks that have the **Flag1** set to **True**.


```vb
Sub DeleteNonEssentialTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 ' Delete nonessential tasks in the active project. 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.Flag1 = True Then T.Delete 
 End If 
 Next T 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]