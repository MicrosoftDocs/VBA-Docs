---
title: Project.OutlineCodes property (Project)
ms.prod: project-server
api_name:
- Project.Project.OutlineCodes
ms.assetid: 400701e8-0114-0819-716f-d79d08a955d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.OutlineCodes property (Project)

Gets an **[OutlineCodes](Project.outlinecodes(object).md)** collection of all outline codes defined for resources and tasks in the project. Read-only **OutlineCodes**.


## Syntax

_expression_. `OutlineCodes`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

Use the local  **OutlineCode[1-10]** properties of resources or tasks to set outline code values in an open project. For example, use the **[OutlineCode4](Project.Task.OutlineCode4.md)** property of the **Task** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]