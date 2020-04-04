---
title: SharingItem.IsConflict property (Outlook)
keywords: vbaol11.chm680
f1_keywords:
- vbaol11.chm680
ms.prod: outlook
api_name:
- Outlook.SharingItem.IsConflict
ms.assetid: 7cf12cb0-71f7-0692-26f0-b20e8a47deed
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.IsConflict property (Outlook)

Returns a **Boolean** that determines if the **[SharingItem](Outlook.SharingItem.md)** is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]