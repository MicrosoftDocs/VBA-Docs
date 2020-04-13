---
title: TaskRequestDeclineItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.Read
ms.assetid: 369c5fe3-2187-46ae-ef68-89734e1296ab
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Remarks

The **Read** event differs from the **[Open](Outlook.TaskRequestDeclineItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]