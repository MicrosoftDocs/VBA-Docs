---
title: MeetingItem.AutoResolvedWinner property (Outlook)
keywords: vbaol11.chm1467
f1_keywords:
- vbaol11.chm1467
ms.prod: outlook
api_name:
- Outlook.MeetingItem.AutoResolvedWinner
ms.assetid: 5a6c9fbb-0f41-9b69-dd41-35ec72e16c7c
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.AutoResolvedWinner property (Outlook)

Returns a **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

_expression_. `AutoResolvedWinner`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](Outlook.Conflicts.Count.md)** of its **[MeetingItem.Conflicts](Outlook.MeetingItem.Conflicts.md)** property greater than zero and if its **AutoResolvedWinner** property is **True**, it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False**, it is a loser in an automatic conflict resolution.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]