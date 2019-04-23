---
title: JournalItem.IsConflict property (Outlook)
keywords: vbaol11.chm1285
f1_keywords:
- vbaol11.chm1285
ms.prod: outlook
api_name:
- Outlook.JournalItem.IsConflict
ms.assetid: 0390d347-959b-0dac-4f8b-7a714c4bdf6e
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.IsConflict property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]