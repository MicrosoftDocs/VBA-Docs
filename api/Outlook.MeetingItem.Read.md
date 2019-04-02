---
title: MeetingItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.Read
ms.assetid: 8a83b213-1afb-7ded-eb67-3e5d21502c5b
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.MeetingItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]