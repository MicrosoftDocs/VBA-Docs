---
title: Group2.Name property (Project)
ms.prod: project-server
api_name:
- Project.Group2.Name
ms.assetid: 27110629-c022-3587-7b9c-c33fbd323a11
ms.date: 06/08/2017
localization_priority: Normal
---


# Group2.Name property (Project)

Gets or sets the name of a **Group2** object. Read/write **String**.


## Syntax

_expression_.**Name**

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