---
title: TaskRequestAcceptItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.Read
ms.assetid: 2a82a5f1-545a-01e4-223f-ca3b31264a4b
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.TaskRequestAcceptItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]