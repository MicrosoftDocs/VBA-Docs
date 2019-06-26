---
title: Project.ProjectStart property (Project)
ms.prod: project-server
api_name:
- Project.Project.ProjectStart
ms.assetid: e29a67b8-fd54-b7ed-3eb0-da4adfa66b6a
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.ProjectStart property (Project)

Gets or sets the start date for a project. Read/write  **Variant**.


## Syntax

_expression_. `ProjectStart`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

Setting  **ProjectStart** value also causes the project to be scheduled from its start date. This has the same effect as setting the **ScheduleFromStart** property to **True**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]