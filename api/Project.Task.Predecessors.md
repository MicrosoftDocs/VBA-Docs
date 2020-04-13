---
title: Task.Predecessors property (Project)
ms.prod: project-server
api_name:
- Project.Task.Predecessors
ms.assetid: 4912eb9f-ad7b-68af-8c3b-c066715c1777
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Predecessors property (Project)

Gets or sets a list of the identification numbers of a task's predecessors. Read/write  **String**.


## Syntax

_expression_. `Predecessors`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

If the predecessors of the specified task have identification numbers of 2 and 10, and the list separator character is the comma, the **Predecessors** property returns "2,10".

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]