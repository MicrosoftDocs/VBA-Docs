---
title: Task.Status property (Project)
ms.prod: project-server
api_name:
- Project.Task.Status
ms.assetid: 4ea3a033-2306-8ae1-4e5e-c0420dcfa3dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Status property (Project)

Gets the status of a specified task. Read-only  **PjStatusType**.


## Syntax

_expression_. `Status`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

The Status property can be one of the following  **[PjStatusType](Project.PjStatusType.md)** constants: **pjComplete**, **pjFutureTask**, **pjLate**, **pjNoData**, or **pjOnSchedule**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]