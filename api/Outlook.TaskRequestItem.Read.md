---
title: TaskRequestItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Read
ms.assetid: 56fc2d07-6d17-874a-0734-db64fa4ccfd6
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

The **Read** event differs from the **[Open](Outlook.TaskRequestItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]