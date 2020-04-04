---
title: TaskRequestItem.AutoResolvedWinner property (Outlook)
keywords: vbaol11.chm1913
f1_keywords:
- vbaol11.chm1913
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.AutoResolvedWinner
ms.assetid: 1e0be069-fcb5-1282-1af4-a80d70506d59
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.AutoResolvedWinner property (Outlook)

Returns a **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

_expression_. `AutoResolvedWinner`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](Outlook.Conflicts.Count.md)** of its **[TaskRequestItem.Conflicts](Outlook.TaskRequestItem.Conflicts.md)** property greater than zero and if its **AutoResolvedWinner** property is **True**, it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False**, it is a loser in an automatic conflict resolution.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]