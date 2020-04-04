---
title: TaskRequestDeclineItem.IsConflict property (Outlook)
keywords: vbaol11.chm1863
f1_keywords:
- vbaol11.chm1863
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.IsConflict
ms.assetid: 41d090c3-18de-84ef-1108-17c7df018182
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.IsConflict property (Outlook)

Returns a **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]