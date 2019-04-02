---
title: TaskItem.StatusOnCompletionRecipients property (Outlook)
keywords: vbaol11.chm1745
f1_keywords:
- vbaol11.chm1745
ms.prod: outlook
api_name:
- Outlook.TaskItem.StatusOnCompletionRecipients
ms.assetid: 9800dcb7-6b12-af4b-0379-25658c946118
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.StatusOnCompletionRecipients property (Outlook)

Returns or sets a semicolon-delimited  **String** of display names for recipients who will receive status upon completion of the task. Read/write.


## Syntax

_expression_. `StatusOnCompletionRecipients`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

This property is calculated from the  **[Recipients](Outlook.TaskItem.Recipients.md)** property. Recipients returned by the **StatusOnCompletionRecipients** property correspond to BCC recipients in the **[Recipients](Outlook.Recipients.md)** collection.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]