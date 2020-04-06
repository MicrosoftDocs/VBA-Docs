---
title: TaskItem.Owner property (Outlook)
keywords: vbaol11.chm1731
f1_keywords:
- vbaol11.chm1731
ms.prod: outlook
api_name:
- Outlook.TaskItem.Owner
ms.assetid: 8af59077-9f4f-2099-fd98-416061447968
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Owner property (Outlook)

Returns or sets a  **String** indicating the owner for the task.


## Syntax

_expression_. `Owner`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

This is a free-form string field. Setting this property to someone other than the current user does not have the effect of delegating the task. Read/write if the task is stored on the Exchange Server public folder. Read-only if it's stored in a user's mailbox or personal folders file.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]