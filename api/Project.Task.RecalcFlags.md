---
title: Task.RecalcFlags property (Project)
ms.service: project-server
api_name:
- Project.Task.RecalcFlags
ms.assetid: d5a5989e-b134-240b-fd37-11f4999e74bc
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Task.RecalcFlags property (Project)

Gets a bit mask, flagging one or more conditions that are driving the task. Read-only **Long**.


## Syntax

_expression_.**RecalcFlags**

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

Use the **[PjRecalcDriverType](Project.PjRecalcDriverType.md)** constants with the return value from the **RecalcFlags** property to determine which specific conditions are driving the task.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]