---
title: Task.Subproject property (Project)
ms.prod: project-server
api_name:
- Project.Task.Subproject
ms.assetid: da054f33-3200-e2bd-4db4-179a30958b98
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Subproject property (Project)

Gets or sets the subproject name for the task. Read/write  **String**.


## Syntax

_expression_. `Subproject`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Example

The following line of code inserts the specified project as a subproject for the task. If the project is not found, it displays a file dialog box with the title  **Cannot find inserted project - C:\Project\MySubProject.mpp**.


```vb
activecell.Task.SubProject = "C:\Project\MySubProject.mpp"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]