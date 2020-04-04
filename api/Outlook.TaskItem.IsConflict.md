---
title: TaskItem.IsConflict property (Outlook)
keywords: vbaol11.chm1764
f1_keywords:
- vbaol11.chm1764
ms.prod: outlook
api_name:
- Outlook.TaskItem.IsConflict
ms.assetid: de713a49-bdc8-363e-4990-cf3535b27981
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.IsConflict property (Outlook)

Returns a **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]