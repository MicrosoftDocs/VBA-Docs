---
title: Task.IsStartValid property (Project)
ms.prod: project-server
ms.assetid: 6e5c90ab-7d7c-1f08-370c-8091d1a55aa6
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.IsStartValid property (Project)

 **True** if the start date of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.


## Syntax

_expression_. `IsStartValid`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The start date of a manually scheduled task can be valid even though the finish date and duration are invalid (empty).

To check the finish date and duration, use the  **[IsFinishValid](Project.task.isfinishvalid.md)** property and the **[IsDurationValid](Project.task.isdurationvalid.md)** property.


## Property value

 **VARIANT**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]