---
title: Project.Tasks property (Project)
ms.prod: project-server
api_name:
- Project.Project.Tasks
ms.assetid: 08bfaadd-9cce-84a2-0ff3-c4b29d9e18cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Tasks property (Project)

Gets a **[Tasks](Project.Task.md)** collection representing the tasks in the project. Read-only **Tasks**.


## Syntax

_expression_. `Tasks`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example displays the name of every task in the active project.


```vb
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveProject.Tasks 
 Names = Names & T.Name & vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]