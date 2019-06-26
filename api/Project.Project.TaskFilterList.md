---
title: Project.TaskFilterList property (Project)
ms.prod: project-server
api_name:
- Project.Project.TaskFilterList
ms.assetid: 303b49c8-cfc3-f4d6-197a-a4dfc130ee85
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.TaskFilterList property (Project)

Gets a  **[List](Project.List.md)** object representing all task filters in the project. Read-only **List**.


## Syntax

_expression_. `TaskFilterList`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example lists all the task filters in the active project.


```vb
Sub SeeAllFilters() 
 
 Dim Temp As Variant 
 Dim TaskFilterNames As String 
 
 For Each Temp In ActiveProject.TaskFilterList 
 TaskFilterNames = TaskFilterNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox TaskFilterNames 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]