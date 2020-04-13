---
title: ContactItem.AutoResolvedWinner property (Outlook)
keywords: vbaol11.chm1088
f1_keywords:
- vbaol11.chm1088
ms.prod: outlook
api_name:
- Outlook.ContactItem.AutoResolvedWinner
ms.assetid: f14ae270-0d3d-5b8c-c85c-9809ba0b82fa
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.AutoResolvedWinner property (Outlook)

Returns a **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.


## Syntax

_expression_. `AutoResolvedWinner`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  **[Conflicts.Count](Outlook.Conflicts.Count.md)** of its **[ContactItem.Conflicts](Outlook.ContactItem.Conflicts.md)** property greater than zero and if its **AutoResolvedWinner** property is **True**, it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its **AutoResolvedWinner** property as **False**, it is a loser in an automatic conflict resolution.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]