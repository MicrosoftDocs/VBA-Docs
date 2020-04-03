---
title: MailItem.TaskCompletedDate property (Outlook)
keywords: vbaol11.chm1395
f1_keywords:
- vbaol11.chm1395
ms.prod: outlook
api_name:
- Outlook.MailItem.TaskCompletedDate
ms.assetid: 4bee35d4-1f1e-0b77-2021-84d4916bef8e
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.TaskCompletedDate property (Outlook)

Returns or sets a  **Date** value that represents the completion date of the task for this **[MailItem](Outlook.MailItem.md)**. Read/write.


## Syntax

_expression_. `TaskCompletedDate`

 _expression_ An expression that returns a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if the **[IsMarkedAsTask](Outlook.MailItem.IsMarkedAsTask.md)** property is set to **False**.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]