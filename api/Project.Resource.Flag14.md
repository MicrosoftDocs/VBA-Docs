---
title: Resource.Flag14 property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Flag14
ms.assetid: 750c51ca-b525-0a8b-c1e1-abb21bee430f
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.Flag14 property (Project)

 **True** if the flag associated with a **Resource** is set. Read/write **Variant**.


## Syntax

_expression_. `Flag14`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


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