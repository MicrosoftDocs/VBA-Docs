---
title: TaskRequestDeclineItem.GetInspector property (Outlook)
keywords: vbaol11.chm1834
f1_keywords:
- vbaol11.chm1834
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.GetInspector
ms.assetid: 8892e56a-275d-b9df-9d9d-bbfd39b98c33
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.GetInspector property (Outlook)

Returns an  **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]