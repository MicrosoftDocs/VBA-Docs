---
title: Task.IsFinishValid property (Project)
ms.prod: project-server
ms.assetid: 13981c95-28fc-7b2f-d8b2-5b235bbe684e
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.IsFinishValid property (Project)

 **True** if the finish date of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.


## Syntax

_expression_. `IsFinishValid`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The finish date of a manually scheduled task can be valid even though the start date and duration are invalid (empty).

To check the start date and duration, use the  **[IsStartValid](Project.task.isstartvalid.md)** property and the **[IsDurationValid](Project.task.isdurationvalid.md)** property.


## Property value

 **VARIANT**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]