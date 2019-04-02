---
title: ContactItem.ToDoTaskOrdinal property (Outlook)
keywords: vbaol11.chm3030
f1_keywords:
- vbaol11.chm3030
ms.prod: outlook
api_name:
- Outlook.ContactItem.ToDoTaskOrdinal
ms.assetid: 080e32ad-b770-42d1-60d0-4eb6271056db
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.ToDoTaskOrdinal property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[ContactItem](Outlook.ContactItem.md)**. Read/write.


## Syntax

_expression_. `ToDoTaskOrdinal`

 _expression_ An expression that returns a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if the **[IsMarkedAsTask](Outlook.ContactItem.IsMarkedAsTask.md)** property is set to **False**.

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](Outlook.ContactItem.TaskStartDate.md)**, **[TaskDueDate](Outlook.ContactItem.TaskDueDate.md)**, or **[TaskCompletedDate](Outlook.ContactItem.TaskCompletedDate.md)** properties.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]