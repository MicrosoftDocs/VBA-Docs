---
title: Project.TaskTableList property (Project)
keywords: vbapj.chm132713
f1_keywords:
- vbapj.chm132713
ms.prod: project-server
api_name:
- Project.Project.TaskTableList
ms.assetid: a36abbcb-db7d-f593-7e5c-df00fd96f010
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.TaskTableList property (Project)

Gets a  **[List](Project.List.md)** object representing all task tables in the project. Read-only **List**.


## Syntax

_expression_. `TaskTableList`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example lists all the task tables in the active project.


```vb
Sub SeeAllTables() 
 
 Dim Temp As Variant 
 Dim TaskTableNames As String 
 
 For Each Temp In ActiveProject.TaskTableList 
 TaskTableNames = TaskTableNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox TaskTableNames 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]