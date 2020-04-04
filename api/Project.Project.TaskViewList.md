---
title: Project.TaskViewList property (Project)
keywords: vbapj.chm132716
f1_keywords:
- vbapj.chm132716
ms.prod: project-server
api_name:
- Project.Project.TaskViewList
ms.assetid: 86d408a2-ed60-fde0-8849-17167d71f6d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.TaskViewList property (Project)

Gets a **[List](Project.List.md)** object representing all task views in the project. Read-only **List**.


## Syntax

_expression_. `TaskViewList`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example lists all the task views in the active project.


```vb
Sub SeeAllViews() 
 
 Dim Temp As Variant 
 Dim TaskViewNames As String 
 
 For Each Temp In ActiveProject.TaskViewList 
 TaskViewNames = TaskViewNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox TaskViewNames 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]