---
title: SharingItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.Read
ms.assetid: 2bcf07e6-e9c1-b3ce-118c-a2c82b48ff5f
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.SharingItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]