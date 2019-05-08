---
title: Task.Flag10 property (Project)
ms.prod: project-server
api_name:
- Project.Task.Flag10
ms.assetid: 68ba415e-8242-986b-dd90-d7ee4dbef58b
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Flag10 property (Project)

Gets or sets the value of a task flag custom field. Read/write  **Variant**.


## Syntax

_expression_. `Flag10`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Example

The following example deletes all the tasks that have the  **Flag1** set to **True**.


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