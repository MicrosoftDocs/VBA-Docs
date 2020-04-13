---
title: TaskRequestAcceptItem.GetInspector property (Outlook)
keywords: vbaol11.chm1785
f1_keywords:
- vbaol11.chm1785
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.GetInspector
ms.assetid: 67239e8b-aa69-c427-3cb5-4a6a1361ed1c
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]