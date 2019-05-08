---
title: Task.StatusManagerName property (Project)
keywords: vbapj.chm132658
f1_keywords:
- vbapj.chm132658
ms.prod: project-server
api_name:
- Project.Task.StatusManagerName
ms.assetid: 4a48ca32-f34b-2225-a687-254c8e3531b1
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.StatusManagerName property (Project)

Gets or sets the GUID of the enterprise resource responsible for accepting or rejecting assignment progress updates for the task. Read/write  **String**.


## Syntax

_expression_. `StatusManagerName`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

The  **StatusManagerName** property is available only in Project Professional. **StatusManagerName** is an empty string for tasks in local projects.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]