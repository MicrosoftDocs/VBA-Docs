---
title: MeetingItem.GetInspector property (Outlook)
keywords: vbaol11.chm1413
f1_keywords:
- vbaol11.chm1413
ms.prod: outlook
api_name:
- Outlook.MeetingItem.GetInspector
ms.assetid: 5e170a6a-6857-ca24-4c14-1e2bc046fd2d
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]