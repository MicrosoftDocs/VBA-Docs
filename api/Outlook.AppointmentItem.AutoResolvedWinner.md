---
title: AppointmentItem.AutoResolvedWinner property (Outlook)
keywords: vbaol11.chm920
f1_keywords:
- vbaol11.chm920
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.AutoResolvedWinner
ms.assetid: d48a7ba9-bb70-9126-98ef-3bdee1f62436
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.AutoResolvedWinner property (Outlook)

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

_expression_. `AutoResolvedWinner`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](Outlook.Conflicts.Count.md)** of its **[AppointmentItem.Conflicts](Outlook.AppointmentItem.Conflicts.md)** property greater than zero and if its **AutoResolvedWinner** property is **True**, it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False**, it is a loser in an automatic conflict resolution.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]