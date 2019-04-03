---
title: Task.Baseline7DurationText property (Project)
keywords: vbapj.chm131531
f1_keywords:
- vbapj.chm131531
ms.prod: project-server
api_name:
- Project.Task.Baseline7DurationText
ms.assetid: 02d9511d-efd7-8641-aa0d-208d6c91420a
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Baseline7DurationText property (Project)

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.


## Syntax

_expression_. `Baseline7DurationText`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The  **Baseline7DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline7DurationText** has any value, you should convert the value to a date for the **Baseline7Duration** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]