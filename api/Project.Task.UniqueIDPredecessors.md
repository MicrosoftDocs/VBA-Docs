---
title: Task.UniqueIDPredecessors property (Project)
keywords: vbapj.chm132772
f1_keywords:
- vbapj.chm132772
ms.prod: project-server
api_name:
- Project.Task.UniqueIDPredecessors
ms.assetid: e6f53dd2-1833-e081-29ee-de734efb9229
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.UniqueIDPredecessors property (Project)

Gets or sets the unique identification (**UniqueID**) numbers of the predecessors of a task, separated by the list separator. Read/write **String**.


## Syntax

_expression_. `UniqueIDPredecessors`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

If a task has two predecessor tasks with the **UniqueID** values of 9 and 10, for example, the **UniqueIDPredecessors** value is "9,10".


> [!NOTE] 
>  **UniqueID** values remain constant within a project and do not necessarily match the task **ID** values that can change with the position of the task in the outline or as tasks are deleted and added.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]