---
title: TaskItem.GetInspector property (Outlook)
keywords: vbaol11.chm1697
f1_keywords:
- vbaol11.chm1697
ms.prod: outlook
api_name:
- Outlook.TaskItem.GetInspector
ms.assetid: 2a2faad7-1030-cdd8-8a8d-8018aad3b667
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]