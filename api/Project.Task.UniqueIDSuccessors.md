---
title: Task.UniqueIDSuccessors property (Project)
keywords: vbapj.chm132773
f1_keywords:
- vbapj.chm132773
ms.prod: project-server
api_name:
- Project.Task.UniqueIDSuccessors
ms.assetid: 2462e6da-8624-62f6-408e-0f50de82096d
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.UniqueIDSuccessors property (Project)

Gets or sets the unique identification (**UniqueID**) numbers of the successors of the task, separated by the list separator. Read/write **String**.


## Syntax

_expression_. `UniqueIDSuccessors`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

If a task has two successor tasks with the **UniqueID** values of 10 and 12, for example, the **UniqueIDSuccessors** value is "10,12".


> [!NOTE] 
>  **UniqueID** values remain constant within a project and do not necessarily match the task **ID** values that can change with the position of the task in the outline or as tasks are deleted and added.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]