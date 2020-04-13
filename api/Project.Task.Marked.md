---
title: Task.Marked property (Project)
ms.prod: project-server
api_name:
- Project.Task.Marked
ms.assetid: b868afee-637f-8725-afdb-3c59ad261e26
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Marked property (Project)

 **True** if the task is marked for further action of some kind. Read/write **Variant**.


## Syntax

_expression_. `Marked`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

For example, you can set the **Marked** property to "Yes", add the **Marked** column to a task view, and filter, format, or edit only those tasks that are marked "Yes". Setting **Marked** to 1 or **True** results in a value of "Yes" for the task in the **Marked** column.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]