---
title: MailItem.ToDoTaskOrdinal property (Outlook)
keywords: vbaol11.chm3038
f1_keywords:
- vbaol11.chm3038
ms.prod: outlook
api_name:
- Outlook.MailItem.ToDoTaskOrdinal
ms.assetid: d1ccb01a-0792-3779-3f94-eb5195a39bb0
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.ToDoTaskOrdinal property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[MailItem](Outlook.MailItem.md)**. Read/write.


## Syntax

_expression_. `ToDoTaskOrdinal`

 _expression_ An expression that returns a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if the **[IsMarkedAsTask](Outlook.MailItem.IsMarkedAsTask.md)** property is set to **False**.

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](Outlook.MailItem.TaskStartDate.md)**, **[TaskDueDate](Outlook.MailItem.TaskDueDate.md)**, or **[TaskCompletedDate](Outlook.MailItem.TaskCompletedDate.md)** properties.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]