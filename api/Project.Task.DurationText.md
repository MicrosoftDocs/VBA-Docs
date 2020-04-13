---
title: Task.DurationText property (Project)
ms.prod: project-server
api_name:
- Project.Task.DurationText
ms.assetid: 4b0bbf0c-13fa-fcab-9940-b3471eb3509b
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.DurationText property (Project)

Gets or sets a string representation of the task duration. Read/write  **String**.


## Syntax

_expression_. `DurationText`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The **DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **DurationText** has any value, you should convert the value to a date for the **Duration** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]