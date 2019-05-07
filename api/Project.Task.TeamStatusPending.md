---
title: Task.TeamStatusPending property (Project)
ms.prod: project-server
api_name:
- Project.Task.TeamStatusPending
ms.assetid: 4c20c56d-d782-5364-0ac8-e19b93f6a887
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.TeamStatusPending property (Project)

 **True** if a response has not been received for at least one progress request message. Read-only **Boolean**.


## Syntax

_expression_. `TeamStatusPending`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

To see whether a team member assigned to the task has responded to an Update Progress request, add the  **TeamStatusPending** field to the task view.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]