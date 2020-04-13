---
title: Task.StartText property (Project)
ms.prod: project-server
api_name:
- Project.Task.StartText
ms.assetid: 32a19317-a16b-c64f-d21f-cdb76d182743
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.StartText property (Project)

Gets or sets a string representation of the task start date. Read/write  **String**.


## Syntax

_expression_. `StartText`

 _expression_ An expression that returns a [Task](./Project.Task.md) object.


## Remarks

The **StartText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **StartText** has any value, you should convert the value to a date for the **Start** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]