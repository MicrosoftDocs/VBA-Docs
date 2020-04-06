---
title: TaskRequestItem.GetInspector property (Outlook)
keywords: vbaol11.chm1883
f1_keywords:
- vbaol11.chm1883
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.GetInspector
ms.assetid: 114a879a-9e5c-5f90-0621-082348dab1df
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]