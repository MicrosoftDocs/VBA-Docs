---
title: TaskRequestItem.IsConflict property (Outlook)
keywords: vbaol11.chm1912
f1_keywords:
- vbaol11.chm1912
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.IsConflict
ms.assetid: d2ab2c17-ba99-1958-38b7-27529cc498e9
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.IsConflict property (Outlook)

Returns a **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]