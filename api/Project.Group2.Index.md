---
title: Group2.Index property (Project)
ms.prod: project-server
api_name:
- Project.Group2.Index
ms.assetid: a7d4ec3e-825b-87c8-d7bb-a61984ba7ace
ms.date: 06/08/2017
localization_priority: Normal
---


# Group2.Index property (Project)

Gets the index of a  **Group2** object in a **ResourceGroups2** collection or **TaskGroups2** collection. Read-only **Long**.


## Syntax

_expression_.**Index**

 _expression_ An expression that returns a [Group2](./Project.Group2.md) object.


## Example

The following example displays the name of each  **Group2** object in the **TaskGroups2** collection in the Immediate window.


```vb
Sub ListTaskGroups() 

 Dim groupIndex As Integer 

 Dim numTaskGroups As Integer 

 

 numTaskGroups = ActiveProject.TaskGroups2.Count 

 

 For groupIndex = 1 To numTaskGroups 

 Debug.Print ActiveProject.TaskGroups2(groupIndex).Name 

 Next groupIndex 

End Sub
```


## See also


[Group2 Object](Project.Group2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]