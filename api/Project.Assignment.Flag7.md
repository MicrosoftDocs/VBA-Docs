---
title: Assignment.Flag7 property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Flag7
ms.assetid: 8613ebea-1029-e66f-cbf9-6ff29d4063a5
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment.Flag7 property (Project)

 **True** if the flag associated with an **Assignment** is set. Read/write **Variant**.


## Syntax

_expression_. `Flag7`

_expression_ A variable that represents an [Assignment](./Project.Assignment.md) object.


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