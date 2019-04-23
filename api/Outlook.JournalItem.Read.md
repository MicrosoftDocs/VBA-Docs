---
title: JournalItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.Read
ms.assetid: 35111126-291b-73b2-2d89-64d950f1c598
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.JournalItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]