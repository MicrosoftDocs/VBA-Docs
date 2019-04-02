---
title: PostItem.ToDoTaskOrdinal property (Outlook)
keywords: vbaol11.chm3042
f1_keywords:
- vbaol11.chm3042
ms.prod: outlook
api_name:
- Outlook.PostItem.ToDoTaskOrdinal
ms.assetid: 58847d68-b956-3d87-6ed2-127801d3fee3
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.ToDoTaskOrdinal property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[PostItem](Outlook.PostItem.md)**. Read/write.


## Syntax

_expression_. `ToDoTaskOrdinal`

 _expression_ An expression that returns a [PostItem](Outlook.PostItem.md) object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if the **[IsMarkedAsTask](Outlook.PostItem.IsMarkedAsTask.md)** property is set to **False**.

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](Outlook.PostItem.TaskStartDate.md)**, **[TaskDueDate](Outlook.PostItem.TaskDueDate.md)**, or **[TaskCompletedDate](Outlook.PostItem.TaskCompletedDate.md)** properties.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]