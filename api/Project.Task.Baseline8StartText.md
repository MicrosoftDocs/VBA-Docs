---
title: Task.Baseline8StartText property (Project)
ms.prod: project-server
api_name:
- Project.Task.Baseline8StartText
ms.assetid: f9ce5373-f49b-e28b-1323-b0ac0896df09
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Baseline8StartText property (Project)

Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.


## Syntax

_expression_. `Baseline8StartText`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The  **Baseline8StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline8StartText** has any value, you should convert the value to a date for the **Baseline8Start** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]