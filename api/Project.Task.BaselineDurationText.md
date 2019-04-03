---
title: Task.BaselineDurationText property (Project)
ms.prod: project-server
api_name:
- Project.Task.BaselineDurationText
ms.assetid: 87307d59-3307-1ee1-82f3-87840d1b4e7a
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.BaselineDurationText property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

_expression_. `BaselineDurationText`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The  **BaselineDurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **BaselineDurationText** has any value, you should convert the value to a date for the **BaselineDuration** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]