---
title: TaskRequestUpdateItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Read
ms.assetid: f324f6b2-dda8-d481-a470-eb660614b6c1
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.TaskRequestUpdateItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]