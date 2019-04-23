---
title: TaskItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Read
ms.assetid: 88e5e300-e036-b511-905c-f0c238c97ade
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.TaskItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]