---
title: Task.IsDurationValid property (Project)
ms.prod: project-server
ms.assetid: 303c5cab-b83a-37b6-c1da-207e91c45a86
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.IsDurationValid property (Project)

 **True** if the duration of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.


## Syntax

_expression_. `IsDurationValid`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

A manually scheduled task must have a valid start date and finish date for the duration to be valid.

To check the start date and finish date, use the **[IsStartValid](Project.task.isstartvalid.md)** property and the **[IsFinishValid](Project.task.isfinishvalid.md)** property.


## Property value

 **VARIANT**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]