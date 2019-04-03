---
title: MeetingItem.IsConflict property (Outlook)
keywords: vbaol11.chm1464
f1_keywords:
- vbaol11.chm1464
ms.prod: outlook
api_name:
- Outlook.MeetingItem.IsConflict
ms.assetid: 1e84c838-06f6-823f-1605-8085d42bb0a0
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.IsConflict property (Outlook)

Returns a  **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]