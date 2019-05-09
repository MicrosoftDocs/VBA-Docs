---
title: Task.OutlineLevel property (Project)
ms.prod: project-server
api_name:
- Project.Task.OutlineLevel
ms.assetid: 7b852e27-bdbc-ee01-4146-c22b929adfa5
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.OutlineLevel property (Project)

Gets the level of the task in the outline hierarchy. Read/write  **Integer**.


## Syntax

_expression_.**OutlineLevel**

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

A task with an outline level of 1 is at the highest level in the outline; there are no summary tasks above it. A task with an outline level of 3 has two summary tasks above it.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]