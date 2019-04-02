---
title: RemoteItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Read
ms.assetid: 78ad2650-7108-f617-6a04-74d7db8db4d7
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.RemoteItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]